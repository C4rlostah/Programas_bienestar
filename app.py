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

load_dotenv()

app = FastAPI(title="Programas Bienestar")
app.add_middleware(SessionMiddleware, secret_key="bienestar_super_secreto_2026")
templates = Jinja2Templates(directory="templates")

DATABASE_URL = os.getenv("DATABASE_URL")

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
            folio_tarjeta TEXT, foto_frente TEXT, foto_reverso TEXT
        )
    ''')
    
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
    registros = cursor.fetchall()
    conn.close()
    return templates.TemplateResponse(request=request, name="registros.html", context={"registros": registros, "user": user})

@app.post("/agregar")
async def agregar_registro(
    request: Request, nombre: str = Form(...), fecha_nacimiento: str = Form(...), seccion: str = Form(...), calle: str = Form(...), 
    num_ext: str = Form(...), num_int: str = Form(""), cp: str = Form(...), fraccionamiento: str = Form(...),
    telefono: str = Form(...), folio_tarjeta: str = Form(...), foto_frente: str = Form(""), foto_reverso: str = Form("")
):
    user = get_current_user(request)
    if not user: return JSONResponse(content={"status": "error", "message": "Sesión expirada"}, status_code=401)

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO beneficiarios (nombre, fecha_nacimiento, seccion, calle, num_ext, num_int, cp, fraccionamiento, telefono, folio_tarjeta, foto_frente, foto_reverso) 
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s) RETURNING id
    ''', (nombre, fecha_nacimiento, seccion, calle, num_ext, num_int, cp, fraccionamiento, telefono, folio_tarjeta, foto_frente, foto_reverso))
    new_id = cursor.fetchone()[0]
    conn.commit()
    conn.close()
    
    registrar_movimiento(user["username"], "Captura", f"Registró al beneficiario: {nombre} (ID: {new_id})")
    return JSONResponse(content={"status": "success", "message": "Registro guardado correctamente"})

@app.post("/editar/{reg_id}")
async def editar_registro(
    request: Request, reg_id: int, nombre: str = Form(...), fecha_nacimiento: str = Form(...), 
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
        "nombre=%s", "fecha_nacimiento=%s", "seccion=%s", "calle=%s", "num_ext=%s", 
        "num_int=%s", "cp=%s", "fraccionamiento=%s", "telefono=%s", "folio_tarjeta=%s"
    ]
    values = [nombre, fecha_nacimiento, seccion, calle, num_ext, num_int, cp, fraccionamiento, telefono, folio_tarjeta]

    if foto_frente:
        updates.append("foto_frente=%s")
        values.append(foto_frente)
    if foto_reverso:
        updates.append("foto_reverso=%s")
        values.append(foto_reverso)

    values.append(reg_id)
    query = f"UPDATE beneficiarios SET {', '.join(updates)} WHERE id=%s"
    
    cursor.execute(query, tuple(values))
    conn.commit()
    conn.close()
    
    registrar_movimiento(user["username"], "Edición", f"Editó datos del beneficiario: {nombre} (ID: {reg_id})")
    return JSONResponse(content={"status": "success"})

@app.post("/eliminar/{reg_id}")
async def eliminar_registro(request: Request, reg_id: int):
    user = get_current_user(request)
    if not user or user["rol"] not in ["admin", "soporte"]:
        return JSONResponse(content={"status": "error", "message": "Acceso denegado"}, status_code=403)
        
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    cursor.execute("SELECT nombre FROM beneficiarios WHERE id = %s", (reg_id,))
    beneficiario = cursor.fetchone()
    
    if beneficiario:
        cursor.execute("DELETE FROM beneficiarios WHERE id = %s", (reg_id,))
        conn.commit()
        registrar_movimiento(user["username"], "Eliminación", f"Eliminó al beneficiario: {beneficiario['nombre']} (ID: {reg_id})")
        
    conn.close()
    return JSONResponse(content={"status": "success"})

def procesar_celda_imagen(worksheet, base64_str, row_idx, col_idx, col_letter):
    if not base64_str:
        return  
        
    if "application/pdf" in base64_str:
        worksheet.cell(row=row_idx, column=col_idx).value = "📄 PDF Adjunto"
        worksheet.cell(row=row_idx, column=col_idx).font = Font(color="0000FF", italic=True)
        return

    try:
        format_str, img_str = base64_str.split(';base64,')
        img_data = base64.b64decode(img_str)
        
        img = PILImage.open(BytesIO(img_data))
        img.thumbnail((60, 40)) 
        
        img_io = BytesIO()
        img.save(img_io, format=img.format)
        img_io.seek(0)
        
        excel_img = ExcelImage(img_io)
        cell_position = f"{col_letter}{row_idx}"
        worksheet.add_image(excel_img, cell_position)
        
    except Exception as e:
        worksheet.cell(row=row_idx, column=col_idx).value = "Error al cargar"

@app.get("/descargar_excel")
async def descargar_excel(request: Request, con_imagenes: str = "true"):
    user = get_current_user(request)
    if not user or user["rol"] not in ["admin", "soporte"]:
        return RedirectResponse(url="/", status_code=303)
    
    incluir_fotos = con_imagenes.lower() == "true"
    
    texto_auditoria = "Descargó la base de datos en Excel (Con Imágenes)" if incluir_fotos else "Descargó la base de datos en Excel (Solo Texto)"
    registrar_movimiento(user["username"], "Exportación", texto_auditoria)
    
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    cursor.execute('''
        SELECT id AS "ID", folio_tarjeta AS "Folio", nombre AS "Nombre", 
               fecha_nacimiento AS "Nacimiento", seccion AS "Sección", 
               calle AS "Calle", num_ext AS "Num_Ext", num_int AS "Num_Int", 
               fraccionamiento AS "Fraccionamiento", cp AS "CP", telefono AS "Teléfono",
               foto_frente, foto_reverso
        FROM beneficiarios
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
    
    # Cambiamos el nombre del archivo dependiendo de la opción elegida
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