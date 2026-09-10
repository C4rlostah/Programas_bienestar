from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse, JSONResponse
from fastapi.templating import Jinja2Templates
import psycopg2
from psycopg2.extras import RealDictCursor
import pandas as pd
from io import BytesIO
import os
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = FastAPI(title="Programas Bienestar")
templates = Jinja2Templates(directory="templates")

# Obtenemos la URL de la base de datos de Neon desde las variables de entorno
DATABASE_URL = os.getenv("DATABASE_URL")
PASSWORD_ADMIN = "bienestar2026"

def get_db_connection():
    return psycopg2.connect(DATABASE_URL)

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS beneficiarios (
            id SERIAL PRIMARY KEY,
            nombre TEXT, fecha_nacimiento TEXT, seccion TEXT,
            calle TEXT, num_ext TEXT, num_int TEXT, cp TEXT, fraccionamiento TEXT, telefono TEXT,
            folio_tarjeta TEXT, foto_frente TEXT, foto_reverso TEXT
        )
    ''')
    conn.commit()
    conn.close()

@app.on_event("startup")
def startup():
    try:
        # Solo intenta crear la tabla si hay una URL configurada
        if DATABASE_URL:
            init_db()
    except Exception as e:
        print("Error conectando a la BD:", e)

@app.get("/", response_class=HTMLResponse)
async def inicio(request: Request):
    return templates.TemplateResponse(request=request, name="index.html")

@app.get("/registros", response_class=HTMLResponse)
async def ver_registros(request: Request):
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    cursor.execute("SELECT * FROM beneficiarios ORDER BY id DESC")
    registros = cursor.fetchall()
    conn.close()
    return templates.TemplateResponse(request=request, name="registros.html", context={"registros": registros})

@app.post("/agregar")
async def agregar_registro(
    nombre: str = Form(...), fecha_nacimiento: str = Form(...), seccion: str = Form(...), calle: str = Form(...), 
    num_ext: str = Form(...), num_int: str = Form(""), cp: str = Form(...), fraccionamiento: str = Form(...),
    telefono: str = Form(...), folio_tarjeta: str = Form(...), foto_frente: str = Form(""), foto_reverso: str = Form("")
):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO beneficiarios (nombre, fecha_nacimiento, seccion, calle, num_ext, num_int, cp, fraccionamiento, telefono, folio_tarjeta, foto_frente, foto_reverso) 
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    ''', (nombre, fecha_nacimiento, seccion, calle, num_ext, num_int, cp, fraccionamiento, telefono, folio_tarjeta, foto_frente, foto_reverso))
    conn.commit()
    conn.close()
    return JSONResponse(content={"status": "success", "message": "Registro guardado correctamente"})

@app.post("/eliminar/{reg_id}")
async def eliminar_registro(reg_id: int, password: str = Form(...)):
    if password != PASSWORD_ADMIN:
        return JSONResponse(content={"status": "error", "message": "Contraseña incorrecta"})
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM beneficiarios WHERE id = %s", (reg_id,))
    conn.commit()
    conn.close()
    return JSONResponse(content={"status": "success"})

@app.post("/editar/{reg_id}")
async def editar_registro(
    reg_id: int, password: str = Form(...), nombre: str = Form(...), fecha_nacimiento: str = Form(...), 
    seccion: str = Form(...), calle: str = Form(...), num_ext: str = Form(...), num_int: str = Form(""),
    cp: str = Form(...), fraccionamiento: str = Form(...), telefono: str = Form(...), folio_tarjeta: str = Form(...)
):
    if password != PASSWORD_ADMIN:
        return JSONResponse(content={"status": "error", "message": "Contraseña incorrecta"})
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE beneficiarios 
        SET nombre=%s, fecha_nacimiento=%s, seccion=%s, calle=%s, num_ext=%s, num_int=%s, cp=%s, fraccionamiento=%s, telefono=%s, folio_tarjeta=%s 
        WHERE id=%s
    ''', (nombre, fecha_nacimiento, seccion, calle, num_ext, num_int, cp, fraccionamiento, telefono, folio_tarjeta, reg_id))
    conn.commit()
    conn.close()
    return JSONResponse(content={"status": "success"})

@app.get("/descargar_excel")
async def descargar_excel():
    conn = get_db_connection()
    df = pd.read_sql_query('SELECT id AS "ID", folio_tarjeta AS "Folio", nombre AS "Nombre", fecha_nacimiento AS "Nacimiento", seccion AS "Sección", calle AS "Calle", num_ext AS "Num_Ext", num_int AS "Num_Int", fraccionamiento AS "Fraccionamiento", cp AS "CP", telefono AS "Teléfono" FROM beneficiarios', conn)
    conn.close()
    
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Beneficiarios')
        worksheet = writer.sheets['Beneficiarios']
        guinda_fill = PatternFill(start_color="691C32", end_color="691C32", fill_type="solid")
        white_font = Font(color="FFFFFF", bold=True, size=11)
        thin_border = Border(left=Side(style='thin', color='DDDDDD'), right=Side(style='thin', color='DDDDDD'), top=Side(style='thin', color='DDDDDD'), bottom=Side(style='thin', color='DDDDDD'))
        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.freeze_panes = 'A2'
        
        for col_num, cell in enumerate(worksheet[1], 1):
            cell.fill = guinda_fill
            cell.font = white_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            
        for idx, col in enumerate(worksheet.columns, 1):
            max_length = 0
            col_letter = get_column_letter(idx)
            for cell in col:
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except: pass
            worksheet.column_dimensions[col_letter].width = min(max_length + 3, 40)
            
    buffer.seek(0)
    return StreamingResponse(buffer, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', headers={'Content-Disposition': 'attachment; filename="Base_Datos_Bienestar.xlsx"'})