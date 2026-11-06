from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware
import psycopg2
from psycopg2.extras import RealDictCursor
import pandas as pd
from io import BytesIO
import os
import hashlib
import uuid
import base64
from datetime import datetime
from dotenv import load_dotenv

from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PILImage

# --- NUEVA LIBRERÍA DE AWS ---
import boto3
from botocore.exceptions import ClientError

load_dotenv()

app = FastAPI(title="Programas Bienestar")
app.add_middleware(SessionMiddleware, secret_key="bienestar_super_secreto_2026")
templates = Jinja2Templates(directory="templates")

DATABASE_URL = os.getenv("DATABASE_URL")

# --- CONFIGURACIÓN DE AWS S3 ---
S3_BUCKET = os.getenv("AWS_BUCKET_NAME")
s3_client = boto3.client(
    's3',
    aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
    aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
    region_name=os.getenv("AWS_REGION", "us-east-1")
)

def subir_a_s3(base64_str, prefijo="expediente"):
    """Recibe un Base64, lo decodifica, lo sube a S3 y devuelve la ruta del archivo."""
    if not base64_str:
        return ""
    # Si no empieza con "data:", significa que ya es una ruta de S3 (para cuando editan sin cambiar foto)
    if not base64_str.startswith("data:"):
        return base64_str
        
    try:
        header, encoded = base64_str.split(";", 1)
        mime_type = header.split(":")[1]
        encoded_data = encoded.split(",", 1)[1]
        
        ext = ".pdf" if "pdf" in mime_type else ".jpg"
        file_name = f"ines/{prefijo}_{uuid.uuid4().hex}{ext}"
        file_bytes = base64.b64decode(encoded_data)
        
        s3_client.put_object(
            Bucket=S3_BUCKET,
            Key=file_name,
            Body=file_bytes,
            ContentType=mime_type
        )
        return file_name # Solo guardamos la llave (ej. ines/exp_123.jpg) en la BD
    except Exception as e:
        print(f"Error subiendo a S3: {e}")
        return ""

def generar_url_firmada(ruta_db):
    """Convierte la ruta guardada en la BD en un enlace temporal seguro de S3."""
    if not ruta_db:
        return ""
    # Si es un Base64 antiguo heredado de la BD vieja, lo devolvemos tal cual para no romper registros pasados
    if ruta_db.startswith("data:"):
        return ruta_db
        
    try:
        # Genera un link seguro que caduca en 1 hora (3600 segundos)
        url = s3_client.generate_presigned_url(
            'get_object',
            Params={'Bucket': S3_BUCKET, 'Key': ruta_db},
            ExpiresIn=3600
        )
        return url
    except Exception as e:
        print(f"Error generando URL firmada: {e}")
        return ""

def hash_password(password: str):
    return hashlib.sha256(password.encode()).hexdigest()

def get_db_connection():
    return psycopg2.connect(DATABASE_URL)

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS beneficiarios (
            id SERIAL PRIMARY KEY, nombre TEXT, fecha_nacimiento TEXT, seccion TEXT,
            calle TEXT, num_ext TEXT, num_int TEXT, cp TEXT, fraccionamiento TEXT, telefono TEXT,
            folio_tarjeta TEXT, foto_frente TEXT, foto_reverso TEXT, fecha_registro TEXT
        )
    ''')
    try:
        cursor.execute("ALTER TABLE beneficiarios ADD COLUMN fecha_registro TEXT")
        cursor.execute("UPDATE beneficiarios SET fecha_registro = %s WHERE fecha_registro IS NULL", (datetime.now().strftime("%Y-%m-%d"),))
    except psycopg2.errors.DuplicateColumn:
        conn.rollback()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS usuarios (
            id SERIAL PRIMARY KEY, username TEXT UNIQUE, password TEXT, rol TEXT, session_token TEXT
        )
    ''')
    try:
        cursor.execute("ALTER TABLE usuarios ADD COLUMN session_token TEXT")
    except psycopg2.errors.DuplicateColumn:
        conn.rollback()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS auditoria (
            id SERIAL PRIMARY KEY, usuario TEXT, accion TEXT, detalles TEXT, 
            fecha TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    cursor.execute("SELECT COUNT(*) FROM usuarios")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO usuarios (username, password, rol) VALUES (%s, %s, %s)", ("admin", hash_password("admin123"), "admin"))
        cursor.execute("INSERT INTO usuarios (username, password, rol) VALUES (%s, %s, %s)", ("soporte", hash_password("soporte123"), "soporte"))
        cursor.execute("INSERT INTO usuarios (username, password, rol) VALUES (%s, %s, %s)", ("captura", hash_password("captura123"), "capturador"))
                       
    conn.commit()
    conn.close()

@app.on_event("startup")
def startup():
    if DATABASE_URL:
        init_db()

def registrar_movimiento(usuario: str, accion: str, detalles: str):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT INTO auditoria (usuario, accion, detalles) VALUES (%s, %s, %s)", (usuario, accion, detalles))
    conn.commit()
    conn.close()

def get_current_user(request: Request):
    user_session = request.session.get("user")
    if not user_session: return None
    
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    cursor.execute("SELECT rol, session_token FROM usuarios WHERE username = %s", (user_session["username"],))
    db_user = cursor.fetchone()
    conn.close()
    
    if not db_user or db_user["session_token"] != user_session.get("token"):
        request.session.clear()
        return None
    
    user_session["rol"] = db_user["rol"]
    return user_session

@app.get("/login", response_class=HTMLResponse)
async def login_get(request: Request):
    return templates.TemplateResponse(request=request, name="login.html")

@app.post("/login")
async def login_post(request: Request, username: str = Form(...), password: str = Form(...)):
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    cursor.execute("SELECT * FROM usuarios WHERE username = %s AND password = %s", (username, hash_password(password)))
    user = cursor.fetchone()
    
    if user:
        new_token = str(uuid.uuid4())
        cursor.execute("UPDATE usuarios SET session_token = %s WHERE id = %s", (new_token, user["id"]))
        conn.commit()
        conn.close()
        
        request.session["user"] = {"username": user["username"], "rol": user["rol"], "token": new_token}
        registrar_movimiento(user["username"], "Inicio de Sesión", "El usuario ingresó al sistema.")
        return RedirectResponse(url="/", status_code=303)
    
    conn.close()
    return templates.TemplateResponse(request=request, name="login.html", context={"error": "Credenciales incorrectas"})

@app.get("/logout")
async def logout(request: Request):
    user = get_current_user(request)
    if user: registrar_movimiento(user["username"], "Cierre de Sesión", "El usuario salió del sistema.")
    request.session.clear()
    return RedirectResponse(url="/login", status_code=303)

@app.get("/", response_class=HTMLResponse)
async def inicio(request: Request):
    user = get_current_user(request)
    if not user: return RedirectResponse(url="/login", status_code=303)
    return templates.TemplateResponse(request=request, name="index.html", context={"user": user})

@app.get("/registros", response_class=HTMLResponse)
async def ver_registros(request: Request):
    user = get_current_user(request)
    if not user or user["rol"] not in ["admin", "soporte"]:
        return RedirectResponse(url="/", status_code=303)
        
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    cursor.execute("SELECT * FROM beneficiarios ORDER BY id DESC")
    registros_db = cursor.fetchall()
    conn.close()

    # Sustituimos la ruta de la BD por la URL firmada temporal de AWS antes de mandarlo al HTML
    registros_procesados = []
    for reg in registros_db:
        reg_dict = dict(reg)
        reg_dict['foto_frente'] = generar_url_firmada(reg_dict['foto_frente'])
        reg_dict['foto_reverso'] = generar_url_firmada(reg_dict['foto_reverso'])
        registros_procesados.append(reg_dict)

    return templates.TemplateResponse(request=request, name="registros.html", context={"registros": registros_procesados, "user": user})

@app.post("/agregar")
async def agregar_registro(
    request: Request, nombre: str = Form(...), fecha_registro: str = Form(...), fecha_nacimiento: str = Form(...), 
    seccion: str = Form(...), calle: str = Form(...), num_ext: str = Form(...), num_int: str = Form(""), 
    cp: str = Form(...), fraccionamiento: str = Form(...), telefono: str = Form(...), folio_tarjeta: str = Form(...), 
    foto_frente: str = Form(""), foto_reverso: str = Form("")
):
    user = get_current_user(request)
    if not user: return JSONResponse(content={"status": "error", "message": "Sesión expirada"}, status_code=401)

    # Subimos a S3 y obtenemos solo la llave de ruta (ej. ines/exp_abc123.jpg)
    s3_frente = subir_a_s3(foto_frente, "frente")
    s3_reverso = subir_a_s3(foto_reverso, "reverso")

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO beneficiarios (nombre, fecha_registro, fecha_nacimiento, seccion, calle, num_ext, num_int, cp, fraccionamiento, telefono, folio_tarjeta, foto_frente, foto_reverso) 
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s) RETURNING id
    ''', (nombre, fecha_registro, fecha_nacimiento, seccion, calle, num_ext, num_int, cp, fraccionamiento, telefono, folio_tarjeta, s3_frente, s3_reverso))
    new_id = cursor.fetchone()[0]
    conn.commit()
    conn.close()
    
    registrar_movimiento(user["username"], "Captura", f"Registró al beneficiario: {nombre} (ID: {new_id}) el día {fecha_registro}")
    return JSONResponse(content={"status": "success", "message": "Registro guardado correctamente"})

@app.post("/editar/{reg_id}")
async def editar_registro(
    request: Request, reg_id: int, nombre: str = Form(...), fecha_registro: str = Form(...), fecha_nacimiento: str = Form(...), 
    seccion: str = Form(...), calle: str = Form(...), num_ext: str = Form(...), num_int: str = Form(""),
    cp: str = Form(...), fraccionamiento: str = Form(...), telefono: str = Form(...), folio_tarjeta: str = Form(...),
    foto_frente: str = Form(""), foto_reverso: str = Form("")
):
    user = get_current_user(request)
    if not user or user["rol"] not in ["admin", "soporte"]:
        return JSONResponse(content={"status": "error", "message": "Acceso denegado"}, status_code=403)
        
    conn = get_db_connection()
    cursor = conn.cursor()
    
    updates = [
        "nombre=%s", "fecha_registro=%s", "fecha_nacimiento=%s", "seccion=%s", "calle=%s", "num_ext=%s", 
        "num_int=%s", "cp=%s", "fraccionamiento=%s", "telefono=%s", "folio_tarjeta=%s"
    ]
    values = [nombre, fecha_registro, fecha_nacimiento, seccion, calle, num_ext, num_int, cp, fraccionamiento, telefono, folio_tarjeta]

    # Si enviaron un Base64 nuevo (el usuario subió otra foto al editar), se sube a S3
    if foto_frente:
        updates.append("foto_frente=%s")
        values.append(subir_a_s3(foto_frente, "frente"))
    if foto_reverso:
        updates.append("foto_reverso=%s")
        values.append(subir_a_s3(foto_reverso, "reverso"))

    values.append(reg_id)
    query = f"UPDATE beneficiarios SET {', '.join(updates)} WHERE id=%s"
    
    cursor.execute(query, tuple(values))
    conn.commit()
    conn.close()
    
    registrar_movimiento(user["username"], "Edición", f"Editó datos del beneficiario: {nombre} (ID: {reg_id})")
    return JSONResponse(content={"status": "success"})

def procesar_celda_imagen(worksheet, valor_bd, row_idx, col_idx, col_letter):
    if not valor_bd: return  

    try:
        # Lógica si el registro es heredado (Base64 directo en BD)
        if valor_bd.startswith("data:"):
            if "application/pdf" in valor_bd:
                worksheet.cell(row=row_idx, column=col_idx).value = "📄 PDF"
                worksheet.cell(row=row_idx, column=col_idx).font = Font(color="0000FF", italic=True)
                return
            img_str = valor_bd.split(';base64,')[1]
            img_data = base64.b64decode(img_str)
            
        # Lógica si el registro está en AWS S3
        else:
            if valor_bd.endswith(".pdf"):
                worksheet.cell(row=row_idx, column=col_idx).value = "📄 PDF en S3"
                worksheet.cell(row=row_idx, column=col_idx).font = Font(color="0000FF", italic=True)
                return
            
            # Descargamos los bytes de la imagen desde S3 en memoria para ponerlos en el Excel
            s3_response = s3_client.get_object(Bucket=S3_BUCKET, Key=valor_bd)
            img_data = s3_response['Body'].read()

        # Inserción en el Excel
        img = PILImage.open(BytesIO(img_data))
        img.thumbnail((60, 40)) 
        img_io = BytesIO()
        img.save(img_io, format=img.format)
        img_io.seek(0)
        
        excel_img = ExcelImage(img_io)
        worksheet.add_image(excel_img, f"{col_letter}{row_idx}")
        
    except Exception as e:
        worksheet.cell(row=row_idx, column=col_idx).value = "Error al cargar"

@app.get("/descargar_excel")
async def descargar_excel(request: Request, con_imagenes: str = "true"):
    user = get_current_user(request)
    if not user or user["rol"] not in ["admin", "soporte"]:
        return RedirectResponse(url="/", status_code=303)
    
    incluir_fotos = con_imagenes.lower() == "true"
    
    texto_auditoria = "Descargó la base de datos en Excel (Con Imágenes desde S3)" if incluir_fotos else "Descargó la base de datos en Excel (Solo Texto)"
    registrar_movimiento(user["username"], "Exportación", texto_auditoria)
    
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    cursor.execute('''
        SELECT id AS "ID", fecha_registro AS "F. Registro", folio_tarjeta AS "Folio", nombre AS "Nombre", 
               fecha_nacimiento AS "Nacimiento", seccion AS "Sección", 
               calle AS "Calle", num_ext AS "Num_Ext", num_int AS "Num_Int", 
               fraccionamiento AS "Fraccionamiento", cp AS "CP", telefono AS "Teléfono",
               foto_frente, foto_reverso
        FROM beneficiarios
        ORDER BY id DESC
    ''')
    registros = cursor.fetchall()
    conn.close()
    
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_display = pd.DataFrame([{k: v for k, v in r.items() if k not in ['foto_frente', 'foto_reverso']} for r in registros])
        df_display.to_excel(writer, index=False, sheet_name='Beneficiarios')
        worksheet = writer.sheets['Beneficiarios']
        
        guinda_fill = PatternFill(start_color="691C32", end_color="691C32", fill_type="solid")
        white_font = Font(color="FFFFFF", bold=True, size=11)
        thin_border = Border(left=Side(style='thin', color='DDDDDD'), right=Side(style='thin', color='DDDDDD'), top=Side(style='thin', color='DDDDDD'), bottom=Side(style='thin', color='DDDDDD'))
        worksheet.freeze_panes = 'A2'
        worksheet.auto_filter.ref = worksheet.dimensions

        if incluir_fotos:
            col_frente_idx = len(df_display.columns) + 1
            col_reverso_idx = len(df_display.columns) + 2
            col_frente_letter = get_column_letter(col_frente_idx)
            col_reverso_letter = get_column_letter(col_reverso_idx)
            
            worksheet.cell(row=1, column=col_frente_idx, value="INE Frente").fill = guinda_fill
            worksheet.cell(row=1, column=col_frente_idx).font = white_font
            worksheet.cell(row=1, column=col_reverso_idx, value="INE Reverso").fill = guinda_fill
            worksheet.cell(row=1, column=col_reverso_idx).font = white_font
            
            worksheet.column_dimensions[col_frente_letter].width = 12
            worksheet.column_dimensions[col_reverso_letter].width = 12

        for col_num, cell in enumerate(worksheet[1], 1):
            cell.fill = guinda_fill
            cell.font = white_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            
        for idx, col in enumerate(worksheet.columns, 1):
            if not incluir_fotos or (idx not in [col_frente_idx, col_reverso_idx]):
                max_length = max((len(str(cell.value)) for cell in col if cell.value), default=0)
                worksheet.column_dimensions[get_column_letter(idx)].width = min(max_length + 3, 40)
        
        for row_idx, registro in enumerate(registros, start=2):
            if incluir_fotos:
                worksheet.row_dimensions[row_idx].height = 35
                limite = col_reverso_idx + 1
            else:
                limite = len(df_display.columns) + 1
                
            for col in range(1, limite):
                worksheet.cell(row=row_idx, column=col).alignment = Alignment(horizontal="center", vertical="center")
            
            if incluir_fotos:
                procesar_celda_imagen(worksheet, registro.get('foto_frente'), row_idx, col_frente_idx, col_frente_letter)
                procesar_celda_imagen(worksheet, registro.get('foto_reverso'), row_idx, col_reverso_idx, col_reverso_letter)
            
    buffer.seek(0)
    
    nombre_archivo = "Base_Datos_Bienestar_CON_FOTOS.xlsx" if incluir_fotos else "Base_Datos_Bienestar_SOLO_TEXTO.xlsx"
    return StreamingResponse(buffer, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', headers={'Content-Disposition': f'attachment; filename="{nombre_archivo}"'})

@app.get("/usuarios", response_class=HTMLResponse)
async def ver_usuarios(request: Request):
    user = get_current_user(request)
    if not user or user["rol"] != "admin": return RedirectResponse(url="/", status_code=303)
        
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    cursor.execute("SELECT id, username, rol, CASE WHEN session_token IS NOT NULL THEN 'En línea' ELSE 'Desconectado' END as estado FROM usuarios ORDER BY id ASC")
    usuarios_db = cursor.fetchall()
    conn.close()
    return templates.TemplateResponse(request=request, name="usuarios.html", context={"user": user, "usuarios_db": usuarios_db})

@app.post("/usuarios/crear")
async def crear_usuario(request: Request, username: str = Form(...), password: str = Form(...), rol: str = Form(...)):
    user = get_current_user(request)
    if not user or user["rol"] != "admin": return JSONResponse(content={"status": "error"}, status_code=403)
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO usuarios (username, password, rol) VALUES (%s, %s, %s)", (username, hash_password(password), rol))
        conn.commit()
        registrar_movimiento(user["username"], "Gestión Personal", f"Creó al usuario '{username}' con rol '{rol}'")
        return JSONResponse(content={"status": "success", "message": "Usuario creado exitosamente"})
    except psycopg2.IntegrityError:
        conn.rollback()
        return JSONResponse(content={"status": "error", "message": "El usuario ya existe"})
    finally:
        conn.close()

@app.post("/usuarios/password/{id}")
async def cambiar_password(request: Request, id: int, new_password: str = Form(...)):
    user = get_current_user(request)
    if not user or user["rol"] != "admin": return JSONResponse(content={"status": "error"}, status_code=403)
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE usuarios SET password = %s, session_token = NULL WHERE id = %s RETURNING username", (hash_password(new_password), id))
    target = cursor.fetchone()[0]
    conn.commit()
    conn.close()
    registrar_movimiento(user["username"], "Seguridad", f"Cambió la contraseña y cerró sesión de '{target}'")
    return JSONResponse(content={"status": "success"})

@app.post("/usuarios/eliminar/{id}")
async def eliminar_usuario(request: Request, id: int):
    user = get_current_user(request)
    if not user or user["rol"] != "admin": return JSONResponse(content={"status": "error"}, status_code=403)
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM usuarios WHERE id = %s RETURNING username", (id,))
    target = cursor.fetchone()[0]
    conn.commit()
    conn.close()
    registrar_movimiento(user["username"], "Gestión Personal", f"Eliminó al usuario '{target}'")
    return JSONResponse(content={"status": "success"})

@app.post("/usuarios/logout/{id}")
async def forzar_logout(request: Request, id: int):
    user = get_current_user(request)
    if not user or user["rol"] != "admin": return JSONResponse(content={"status": "error"}, status_code=403)
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE usuarios SET session_token = NULL WHERE id = %s RETURNING username", (id,))
    target = cursor.fetchone()[0]
    conn.commit()
    conn.close()
    registrar_movimiento(user["username"], "Seguridad", f"Forzó el cierre de sesión de '{target}'")
    return JSONResponse(content={"status": "success"})

@app.get("/auditoria", response_class=HTMLResponse)
async def ver_auditoria(request: Request):
    user = get_current_user(request)
    if not user or user["rol"] != "admin": return RedirectResponse(url="/", status_code=303)
        
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    cursor.execute("SELECT * FROM auditoria ORDER BY id DESC LIMIT 200") 
    logs = cursor.fetchall()
    conn.close()
    return templates.TemplateResponse(request=request, name="auditoria.html", context={"user": user, "logs": logs})