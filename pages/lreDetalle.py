"""
Sistema de Gestión de Parámetros Previsionales
- Parámetros Mensuales
- Instituciones AFP (con validación Previred)
Desarrollado con PyQt6 + SQLite
"""

import sys
import sqlite3
import os
import re
import json
from datetime import datetime
from openpyxl import load_workbook
import fitz

# ── Librerías lectura liquidaciones ───────────────────────────────────────────
try:
    import pdfplumber
    _PDFPLUMBER_OK = True
except ImportError:
    _PDFPLUMBER_OK = False

try:
    from pypdf import PdfReader
    _PYPDF_OK = True
except ImportError:
    _PYPDF_OK = False

try:
    import pytesseract
    from PIL import Image
    import io
    _TESSERACT_OK = True
except ImportError:
    _TESSERACT_OK = False

# Ruta del ejecutable Tesseract en Windows (ajustar si se instaló en otra ruta)
_TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

from PyQt6.QtWidgets import (
    QScrollArea, QComboBox,
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QPushButton, QLabel, QLineEdit,
    QDoubleSpinBox, QSpinBox, QMessageBox, QDialog,
    QDialogButtonBox, QHeaderView, QFrame, QStatusBar,
    QGroupBox, QGridLayout, QTabWidget, QTextEdit, QFileDialog,
    QProgressBar, QSplitter, QCheckBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QColor

# ─────────────────────────────────────────────
# RUTAS
# ─────────────────────────────────────────────

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
DB_PATH     = os.path.join(BASE_DIR, "parametros_mensuales.db")
XL_PARAM    = os.path.join(BASE_DIR, "parametrosMesuales.xlsx")
XL_AFP      = os.path.join(BASE_DIR, "inst_afp.xlsx")
XL_PREVIRED = os.path.join(BASE_DIR, "afp_previred.xlsx")

# ─────────────────────────────────────────────
# ESQUEMAS
# ─────────────────────────────────────────────

PARAM_COLUMNS = [
    ("mes_proc",            "TEXT PRIMARY KEY"),
    ("uf_mes",              "REAL"),
    ("tope_imp_uf_afp",     "REAL"),
    ("tope_imp_pesos_afp",  "REAL"),
    ("tope_ces_uf",         "REAL"),
    ("tope_ces_pesos",      "REAL"),
    ("sis",                 "REAL"),
    ("factor_sis",          "REAL"),
    ("tope_salud_uf",       "REAL"),
    ("tope_salud_pesos",    "REAL"),
    ("imm",                 "REAL"),
    ("tope_gratif",         "REAL"),
    ("monto_utm",           "REAL"),
    ("ult_dia_mes",         "INTEGER"),
    ("aporte_ccaf",         "REAL"),
    ("aporte_fonasa",       "REAL"),
    ("formato_fecha",       "TEXT"),
    ("aporte_afp",          "REAL"),
    ("seg_social_exp_vida", "REAL"),
]

PARAM_HEADERS = [
    "Mes Proceso", "UF Mes", "Tope Imp. UF AFP", "Tope Imp. $ AFP",
    "Tope Ces. UF", "Tope Ces. $", "SIS (%)", "Factor SIS",
    "Tope Salud UF", "Tope Salud $", "IMM", "Tope Gratif.",
    "UTM", "Últ. Día", "Aporte CCAF", "Aporte FONASA",
    "Fecha", "Aporte AFP", "Seg. Social"
]

EMPRESAS_HEADERS = ["ID Empresa","RUT","Razon Social","Nom. Fantasia","Region","Comuna","Ciudad","Direccion","Telefono","Email Rep.","CCAF","Mutual","Cot. Mutual","Cod. Act. Econ.","RUT Rep. Legal","Nombre Rep. Legal","Email Empresa"]

CAJAS_HEADERS   = ["ID Inst.", "Clasif.", "Nombre Institucion", "Doc. Identidad", "Cod. Equiv.", "Valor", "Valor 2", "Valor 3", "Dato Adicional"]
MUTUALES_HEADERS = ["ID Inst.", "Clasif.", "Nombre Institucion", "Doc. Identidad", "Cod. Equiv.", "Valor", "Valor 2", "Valor 3", "Dato Adicional"]
APV_HEADERS = ["ID APV", "Clasificacion", "Nombre Institucion APV", "RUT", "Cod. Previred"]

SALUD_HEADERS = ["ID Institucion", "Clasificacion", "Nombre Institucion", "RUT", "Equiv. Previred"]

PREVIRED_SALUD = {"00":"Sin Isapre","01":"Banmedica","02":"Consalud","03":"VidaTres","04":"Colmena","05":"Cruz Blanca","07":"Fonasa","10":"Nueva Masvida","11":"Isalud","12":"Fundacion","25":"Cruz del Norte","28":"Esencial"}

AFP_COLUMNS = [
    ("id_afp",      "TEXT PRIMARY KEY"),
    ("clas_afp",    "TEXT"),
    ("nombre_afp",  "TEXT"),
    ("rut_afp",     "TEXT"),
    ("codPrev_afp", "TEXT"),
    ("cot_afp",     "REAL"),
    ("observ_afp",  "TEXT"),
]

AFP_HEADERS = ["ID AFP", "Clasificación", "Nombre AFP", "RUT", "Cód. Previred", "Cotización (%)", "Observaciones"]

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def normalizar_cod(cod) -> str:
    """Normaliza un código Previred a mínimo 2 caracteres con cero a la izquierda."""
    s = str(cod).strip()
    if s.isdigit():
        return s.zfill(2)
    return s


def ult_dia_para_mes(mes_proc: str) -> int:
    try:
        mm = mes_proc[-2:]
        return 31 if mm in ("01","03","05","07","08","10","12") else 30
    except Exception:
        return 30


def calcular_campos(uf_mes, tope_imp_uf_afp, tope_ces_uf, sis, aporte_ccaf, mes_proc):
    tope_imp_pesos_afp = round(uf_mes * tope_imp_uf_afp) if (uf_mes and tope_imp_uf_afp) else 0
    tope_ces_pesos     = round(uf_mes * tope_ces_uf)     if (uf_mes and tope_ces_uf)     else 0
    factor_sis         = round(sis / 100, 4)             if sis                          else 0
    tope_salud_uf      = round(tope_imp_uf_afp * 0.07, 3) if tope_imp_uf_afp             else 0
    tope_salud_pesos   = round(tope_salud_uf * uf_mes)   if (tope_salud_uf and uf_mes)   else 0
    ult_dia_mes        = ult_dia_para_mes(mes_proc)       if mes_proc                     else 30
    aporte_fonasa      = round(7 - aporte_ccaf, 2)        if aporte_ccaf is not None      else 0
    return {
        "tope_imp_pesos_afp": tope_imp_pesos_afp,
        "tope_ces_pesos":     tope_ces_pesos,
        "factor_sis":         factor_sis,
        "tope_salud_uf":      tope_salud_uf,
        "tope_salud_pesos":   tope_salud_pesos,
        "ult_dia_mes":        ult_dia_mes,
        "aporte_fonasa":      aporte_fonasa,
    }


def siguiente_mes(mes_proc: str) -> str:
    try:
        y, m = int(mes_proc[:4]), int(mes_proc[-2:])
        m += 1
        if m > 12:
            m, y = 1, y + 1
        return f"{y:04d}-{m:02d}"
    except Exception:
        return ""



# ─────────────────────────────────────────────
# EXTRACCIÓN DESDE PDF
# ─────────────────────────────────────────────

MESES_PDF = {
    'enero':'01','febrero':'02','marzo':'03','abril':'04',
    'mayo':'05','junio':'06','julio':'07','agosto':'08',
    'septiembre':'09','octubre':'10','noviembre':'11','diciembre':'12'
}

def extraer_mes_anio_pdf(nombre_archivo):
    """Extrae AAAA-MM desde el nombre del archivo PDF."""
    nombre = os.path.basename(nombre_archivo).lower().replace('.pdf','')
    partes = nombre.split('-')
    anio = None
    mes  = None
    for parte in partes:
        if parte in MESES_PDF:
            mes = MESES_PDF[parte]
        if parte.isdigit() and len(parte) == 4:
            anio = parte
    if mes and anio:
        return f"{anio}-{mes}"
    return None

def limpiar_num_pdf(s):
    """Convierte texto con $ y puntos a float."""
    return float(s.replace('$','').replace(' ','').replace('.','').replace(',','.').strip())

def extraer_datos_pdf(ruta_pdf):
    """Extrae los datos previsionales desde el PDF de Indicadores Previred."""
    try:
        doc    = fitz.open(ruta_pdf)
        lineas = "".join(p.get_text() for p in doc).split('\n')
        datos  = {}

        # UF del mes (línea 1)
        datos['uf_mes'] = limpiar_num_pdf(lineas[1])

        # Tope imponible AFP en pesos (línea 3)
        datos['tope_imp_pesos_afp'] = int(limpiar_num_pdf(lineas[3]))

        # IMM - Trab. Dependientes e Independientes (línea 4)
        datos['imm'] = int(limpiar_num_pdf(lineas[4]))

        # Tope cesantía en pesos (línea 7)
        datos['tope_ces_pesos'] = int(limpiar_num_pdf(lineas[7]))

        # Seg Social Expectativa de Vida (línea 11)
        datos['seg_social_exp_vida'] = float(lineas[11].replace('%','').replace(',','.').strip())

        # Aporte AFP cargo empleador (línea 20)
        datos['aporte_afp'] = float(lineas[20].replace('%','').replace(',','.').strip())

        # SIS (línea 28)
        datos['sis'] = float(lineas[28].replace('%','').replace(',','.').strip())

        # Aporte CCAF (línea 56)
        datos['aporte_ccaf'] = float(lineas[56].replace('%','').replace(' ','').replace('R.I.','').replace(',','.').strip())

        # UTM (línea 89)
        datos['monto_utm'] = int(limpiar_num_pdf(lineas[89]))

        # Campos calculados
        datos['tope_gratif']    = round((datos['imm'] * 4.75) / 12)
        datos['tope_imp_uf_afp'] = round(datos['tope_imp_pesos_afp'] / datos['uf_mes'], 2)
        datos['tope_ces_uf']    = round(datos['tope_ces_pesos'] / datos['uf_mes'], 2)

        return datos, None
    except Exception as e:
        return None, str(e)


# ─────────────────────────────────────────────
# EXPORTAR A EXCEL
# ─────────────────────────────────────────────

def exportar_parametros_excel(ruta_destino):
    """Exporta todos los registros de parámetros mensuales a Excel con formato profesional."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    HDR_COLOR = "1e3a5f"
    ALT_COLOR  = "EFF6FF"
    thin   = Side(style="thin", color="C8D4E0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = Workbook()
    ws = wb.active
    ws.title = "Parámetros Mensuales"

    # Título principal
    ws.merge_cells("A1:S1")
    ws["A1"].value = "Parámetros Mensuales Previsionales"
    ws["A1"].font  = Font(bold=True, color="FFFFFF", size=14, name="Arial")
    ws["A1"].fill  = PatternFill("solid", start_color=HDR_COLOR)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Encabezados
    headers = [
        "Mes Proceso", "UF Mes", "Tope Imp. UF AFP", "Tope Imp. $ AFP",
        "Tope Ces. UF", "Tope Ces. $", "SIS (%)", "Factor SIS",
        "Tope Salud UF", "Tope Salud $", "IMM", "Tope Gratif.",
        "UTM", "Últ. Día", "Aporte CCAF", "Aporte FONASA",
        "Fecha", "Aporte AFP", "Seg. Social"
    ]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=h)
        cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        cell.fill      = PatternFill("solid", start_color="2563eb")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = border
    ws.row_dimensions[2].height = 32

    # Datos
    rows = param_fetch_all()
    for r_idx, row in enumerate(rows):
        alt = r_idx % 2 == 0
        for c_idx, val in enumerate(row):
            cell = ws.cell(row=r_idx+3, column=c_idx+1)
            if val is None:
                cell.value = ""
            elif c_idx in (0, 16):
                cell.value = str(val)
            elif c_idx == 13:
                cell.value = int(val)
            else:
                try:
                    cell.value = float(val)
                except:
                    cell.value = str(val)
            cell.font      = Font(name="Arial", size=10)
            cell.border    = border
            cell.alignment = Alignment(
                horizontal="center" if c_idx in (0, 13, 16) else "right",
                vertical="center"
            )
            if c_idx not in (0, 13, 16) and isinstance(cell.value, float):
                if c_idx in (3, 5, 9, 10, 11, 12):
                    cell.number_format = "#,##0"
                else:
                    cell.number_format = "#,##0.0000"
            if alt:
                cell.fill = PatternFill("solid", start_color=ALT_COLOR)

    # Anchos de columna
    anchos = [12,12,16,16,12,14,10,12,14,14,12,14,12,10,12,14,12,12,12]
    letras = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S"]
    for letra, w in zip(letras, anchos):
        ws.column_dimensions[letra].width = w

    # Congelar fila de encabezados
    ws.freeze_panes = "A3"

    wb.save(ruta_destino)
    return len(rows)

# ─────────────────────────────────────────────
# BASE DE DATOS
# ─────────────────────────────────────────────

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS regiones (num_region TEXT PRIMARY KEY, nombre_region TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS comunas_ciudades (id INTEGER PRIMARY KEY AUTOINCREMENT, num_region TEXT, nombre_region TEXT, provincia TEXT, comuna TEXT, ciudad TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS empresas (id_empresa TEXT PRIMARY KEY, rut_emp TEXT, raz_soc TEXT, nomb_fant TEXT, region TEXT, comuna TEXT, ciudad TEXT, direcc_emp TEXT, fono TEXT, email_repleg TEXT, afp TEXT, salud TEXT, ccaf_emp TEXT, mutual_emp TEXT, cot_mut TEXT, cod_ae TEXT, rut_repleg TEXT, nom_repleg TEXT, email_emp TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS inst_cajas (
        id_inst    TEXT PRIMARY KEY, clasif TEXT, nombre_inst TEXT,
        doc_ident  TEXT, cod_equiv TEXT, valor TEXT, valor2 TEXT,
        valor3 TEXT, dato_adic TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS inst_mutuales (
        id_inst    TEXT PRIMARY KEY, clasif TEXT, nombre_inst TEXT,
        doc_ident  TEXT, cod_equiv TEXT, valor TEXT, valor2 TEXT,
        valor3 TEXT, dato_adic TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS inst_apv (id_apv TEXT PRIMARY KEY, clasif_apv TEXT, nombre_apv TEXT, rut_apv TEXT, codprevi_apv TEXT)""")

    col_defs = ", ".join(f"{c} {t}" for c, t in PARAM_COLUMNS)
    cur.execute(f"CREATE TABLE IF NOT EXISTS parametros ({col_defs})")

    afp_defs = ", ".join(f"{c} {t}" for c, t in AFP_COLUMNS)
    cur.execute(f"CREATE TABLE IF NOT EXISTS inst_afp ({afp_defs})")

    cur.execute("""CREATE TABLE IF NOT EXISTS afp_previred (
        codPrev_afp TEXT PRIMARY KEY,
        nombre_afp  TEXT
    )""")

    conn.commit()

    cur.execute("SELECT COUNT(*) FROM parametros")
    if cur.fetchone()[0] == 0:
        _import_parametros(cur)

    cur.execute("SELECT COUNT(*) FROM inst_afp")
    if cur.fetchone()[0] == 0:
        _import_afp(cur)

    cur.execute("SELECT COUNT(*) FROM afp_previred")
    if cur.fetchone()[0] == 0:
        _import_previred(cur)

    conn.commit()
    conn.close()


def _import_parametros(cur):
    if not os.path.exists(XL_PARAM):
        return
    try:
        wb   = load_workbook(XL_PARAM, read_only=True, data_only=True)
        ws   = wb["Hoja2"]
        rows = [r for r in ws.iter_rows(values_only=True) if any(v is not None for v in r)]
        col_names = [c[0] for c in PARAM_COLUMNS]
        for row in rows[1:]:
            mes = row[0]
            if not mes:
                continue
            uf, tiuf, tcuf, sis, ccaf = row[1], row[2], row[4], row[6], row[14]
            calc = calcular_campos(uf or 0, tiuf or 0, tcuf or 0, sis or 0, ccaf or 0, str(mes))
            ff = row[16]
            if isinstance(ff, datetime):
                ff = ff.strftime("%d-%m-%Y")
            elif ff:
                ff = str(ff)
            values = [
                str(mes), uf, tiuf, calc["tope_imp_pesos_afp"],
                tcuf, calc["tope_ces_pesos"], sis, calc["factor_sis"],
                calc["tope_salud_uf"], calc["tope_salud_pesos"],
                row[10], row[11], row[12], calc["ult_dia_mes"],
                ccaf, calc["aporte_fonasa"], ff, row[17], row[18],
            ]
            ph = ", ".join("?" * len(col_names))
            cur.execute(f"INSERT OR IGNORE INTO parametros ({', '.join(col_names)}) VALUES ({ph})", values)
    except Exception as e:
        print(f"Error importando parámetros: {e}")


def _import_afp(cur):
    if not os.path.exists(XL_AFP):
        return
    try:
        wb   = load_workbook(XL_AFP, read_only=True, data_only=True)
        ws   = wb["Listado"]
        rows = [r for r in ws.iter_rows(values_only=True) if any(v is not None for v in r)]
        col_names = [c[0] for c in AFP_COLUMNS]
        for row in rows[1:]:
            if not row[0]:
                continue
            # Normalizar codPrev_afp al importar
            cod = normalizar_cod(row[4]) if row[4] is not None else ""
            values = [
                str(row[0]).strip(),
                str(row[1]).strip() if row[1] else "",
                str(row[2]).strip() if row[2] else "",
                str(row[3]).strip() if row[3] else "",
                cod,
                float(row[5]) if row[5] else 0.0,
                str(row[6]).strip() if row[6] else "",
            ]
            ph = ", ".join("?" * len(col_names))
            cur.execute(f"INSERT OR IGNORE INTO inst_afp ({', '.join(col_names)}) VALUES ({ph})", values)
    except Exception as e:
        print(f"Error importando AFPs: {e}")


def _import_previred(cur):
    if not os.path.exists(XL_PREVIRED):
        return
    try:
        wb   = load_workbook(XL_PREVIRED, read_only=True, data_only=True)
        ws   = wb["Hoja1"]
        rows = [r for r in ws.iter_rows(values_only=True) if any(v is not None for v in r)]
        for row in rows[1:]:
            if row[0] is None:
                continue
            # Siempre guardar con mínimo 2 caracteres
            cod = normalizar_cod(row[0])
            cur.execute(
                "INSERT OR IGNORE INTO afp_previred (codPrev_afp, nombre_afp) VALUES (?, ?)",
                (cod, str(row[1]).strip())
            )
    except Exception as e:
        print(f"Error importando Previred: {e}")


# ─── CRUD Parámetros ───────────────────────

def param_fetch_all(search=""):
    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()
    if search:
        cur.execute("SELECT * FROM parametros WHERE mes_proc LIKE ? ORDER BY mes_proc DESC", (f"%{search}%",))
    else:
        cur.execute("SELECT * FROM parametros ORDER BY mes_proc DESC")
    rows = cur.fetchall(); conn.close(); return rows

def param_get_last_mes():
    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()
    cur.execute("SELECT mes_proc FROM parametros ORDER BY mes_proc DESC LIMIT 1")
    row = cur.fetchone(); conn.close()
    return row[0] if row else None

def param_get(mes_proc):
    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()
    cur.execute("SELECT * FROM parametros WHERE mes_proc = ?", (mes_proc,))
    row = cur.fetchone(); conn.close(); return row

def param_insert(values):
    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()
    col_names = [c[0] for c in PARAM_COLUMNS]
    ph = ", ".join("?" * len(col_names))
    cur.execute(f"INSERT INTO parametros ({', '.join(col_names)}) VALUES ({ph})", values)
    conn.commit(); conn.close()

def param_update(mes_proc, values):
    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()
    col_names = [c[0] for c in PARAM_COLUMNS if c[0] != "mes_proc"]
    set_clause = ", ".join(f"{c} = ?" for c in col_names)
    cur.execute(f"UPDATE parametros SET {set_clause} WHERE mes_proc = ?", values + [mes_proc])
    conn.commit(); conn.close()

def param_delete(mes_proc):
    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()
    cur.execute("DELETE FROM parametros WHERE mes_proc = ?", (mes_proc,))
    conn.commit(); conn.close()


# ─── CRUD AFP ──────────────────────────────


# ── GEO ──────────────────────────────────────────────────────
def regiones_get_all():
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    cur.execute("SELECT num_region, nombre_region FROM regiones ORDER BY CAST(num_region AS INTEGER)")
    rows=cur.fetchall();conn.close();return rows

def comunas_get_by_region(num_region):
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    cur.execute("SELECT DISTINCT comuna FROM comunas_ciudades WHERE num_region=? ORDER BY comuna",(num_region,))
    rows=cur.fetchall();conn.close();return [r[0] for r in rows]

def ciudades_get_by_comuna(num_region,comuna):
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    cur.execute("SELECT DISTINCT ciudad FROM comunas_ciudades WHERE num_region=? AND comuna=? ORDER BY ciudad",(num_region,comuna))
    rows=cur.fetchall();conn.close();return [r[0] for r in rows]

# ── CRUD empresas ─────────────────────────────────────────────
def empresas_fetch_all(search=""):
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    if search:
        cur.execute("SELECT * FROM empresas WHERE raz_soc LIKE ? OR rut_emp LIKE ? OR id_empresa LIKE ? ORDER BY raz_soc",(f"%{search}%",f"%{search}%",f"%{search}%"))
    else:
        cur.execute("SELECT * FROM empresas ORDER BY raz_soc")
    rows=cur.fetchall();conn.close();return rows

def empresas_get(id_empresa):
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    cur.execute("SELECT * FROM empresas WHERE id_empresa=?",(id_empresa,))
    row=cur.fetchone();conn.close();return row

def empresas_insert(v):
    conn=sqlite3.connect(DB_PATH)
    conn.execute("INSERT INTO empresas VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",v)
    conn.commit();conn.close()

def empresas_update(id_empresa,v):
    conn=sqlite3.connect(DB_PATH)
    conn.execute("UPDATE empresas SET rut_emp=?,raz_soc=?,nomb_fant=?,region=?,comuna=?,ciudad=?,direcc_emp=?,fono=?,email_repleg=?,ccaf_emp=?,mutual_emp=?,cot_mut=?,cod_ae=?,rut_repleg=?,nom_repleg=?,email_emp=? WHERE id_empresa=?",v+[id_empresa])
    conn.commit();conn.close()

def empresas_delete(id_empresa):
    conn=sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM empresas WHERE id_empresa=?",(id_empresa,))
    conn.commit();conn.close()

# ── CRUD inst_cajas ──────────────────────────────────────────
def cajas_fetch_all(search=""):
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    if search:
        cur.execute("SELECT * FROM inst_cajas WHERE nombre_inst LIKE ? OR id_inst LIKE ? ORDER BY nombre_inst",(f"%{search}%",f"%{search}%"))
    else:
        cur.execute("SELECT * FROM inst_cajas ORDER BY nombre_inst")
    rows=cur.fetchall();conn.close();return rows

def cajas_get(id_inst):
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    cur.execute("SELECT * FROM inst_cajas WHERE id_inst=?",(id_inst,))
    row=cur.fetchone();conn.close();return row

def cajas_insert(v):
    conn=sqlite3.connect(DB_PATH)
    conn.execute("INSERT INTO inst_cajas VALUES(?,?,?,?,?,?,?,?,?)",v)
    conn.commit();conn.close()

def cajas_update(id_inst,v):
    conn=sqlite3.connect(DB_PATH)
    conn.execute("UPDATE inst_cajas SET clasif=?,nombre_inst=?,doc_ident=?,cod_equiv=?,valor=?,valor2=?,valor3=?,dato_adic=? WHERE id_inst=?",v+[id_inst])
    conn.commit();conn.close()

def cajas_delete(id_inst):
    conn=sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM inst_cajas WHERE id_inst=?",(id_inst,))
    conn.commit();conn.close()

# ── CRUD inst_mutuales ────────────────────────────────────────
def mutuales_fetch_all(search=""):
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    if search:
        cur.execute("SELECT * FROM inst_mutuales WHERE nombre_inst LIKE ? OR id_inst LIKE ? ORDER BY nombre_inst",(f"%{search}%",f"%{search}%"))
    else:
        cur.execute("SELECT * FROM inst_mutuales ORDER BY nombre_inst")
    rows=cur.fetchall();conn.close();return rows

def mutuales_get(id_inst):
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    cur.execute("SELECT * FROM inst_mutuales WHERE id_inst=?",(id_inst,))
    row=cur.fetchone();conn.close();return row

def mutuales_insert(v):
    conn=sqlite3.connect(DB_PATH)
    conn.execute("INSERT INTO inst_mutuales VALUES(?,?,?,?,?,?,?,?,?)",v)
    conn.commit();conn.close()

def mutuales_update(id_inst,v):
    conn=sqlite3.connect(DB_PATH)
    conn.execute("UPDATE inst_mutuales SET clasif=?,nombre_inst=?,doc_ident=?,cod_equiv=?,valor=?,valor2=?,valor3=?,dato_adic=? WHERE id_inst=?",v+[id_inst])
    conn.commit();conn.close()

def mutuales_delete(id_inst):
    conn=sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM inst_mutuales WHERE id_inst=?",(id_inst,))
    conn.commit();conn.close()

def apv_fetch_all(search=""):
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    if search:
        cur.execute("SELECT * FROM inst_apv WHERE nombre_apv LIKE ? OR id_apv LIKE ? ORDER BY nombre_apv",(f"%{search}%",f"%{search}%"))
    else:
        cur.execute("SELECT * FROM inst_apv ORDER BY nombre_apv")
    rows=cur.fetchall();conn.close();return rows

def apv_get(id_apv):
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    cur.execute("SELECT * FROM inst_apv WHERE id_apv=?",(id_apv,))
    row=cur.fetchone();conn.close();return row

def apv_insert(v):
    conn=sqlite3.connect(DB_PATH)
    conn.execute("INSERT INTO inst_apv VALUES(?,?,?,?,?)",v)
    conn.commit();conn.close()

def apv_update(id_apv,v):
    conn=sqlite3.connect(DB_PATH)
    conn.execute("UPDATE inst_apv SET clasif_apv=?,nombre_apv=?,rut_apv=?,codprevi_apv=? WHERE id_apv=?",v+[id_apv])
    conn.commit();conn.close()

def apv_delete(id_apv):
    conn=sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM inst_apv WHERE id_apv=?",(id_apv,))
    conn.commit();conn.close()

def salud_fetch_all(search=""):
    conn=sqlite3.connect(DB_PATH)
    cur=conn.cursor()
    if search:
        cur.execute("SELECT * FROM inst_salud WHERE nombre_inst LIKE ? OR id_inst LIKE ? ORDER BY nombre_inst",(f"%{search}%",f"%{search}%"))
    else:
        cur.execute("SELECT * FROM inst_salud ORDER BY nombre_inst")
    rows=cur.fetchall();conn.close();return rows

def salud_get(id_inst):
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    cur.execute("SELECT * FROM inst_salud WHERE id_inst=?",(id_inst,))
    row=cur.fetchone();conn.close();return row

def salud_insert(v):
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    cur.execute("INSERT INTO inst_salud VALUES(?,?,?,?,?)",v)
    conn.commit();conn.close()

def salud_update(id_inst,v):
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    cur.execute("UPDATE inst_salud SET clasif=?,nombre_inst=?,rut_inst=?,equiv_previred=? WHERE id_inst=?",v+[id_inst])
    conn.commit();conn.close()

def salud_delete(id_inst):
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    cur.execute("DELETE FROM inst_salud WHERE id_inst=?",(id_inst,))
    conn.commit();conn.close()

def salud_get_by_equiv(equiv,exclude_id=None):
    conn=sqlite3.connect(DB_PATH);cur=conn.cursor()
    if exclude_id:
        cur.execute("SELECT * FROM inst_salud WHERE equiv_previred=? AND id_inst!=?",(equiv,exclude_id))
    else:
        cur.execute("SELECT * FROM inst_salud WHERE equiv_previred=?",(equiv,))
    row=cur.fetchone();conn.close();return row

def afp_fetch_all(search=""):
    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()
    if search:
        cur.execute(
            "SELECT * FROM inst_afp WHERE nombre_afp LIKE ? OR id_afp LIKE ? ORDER BY nombre_afp",
            (f"%{search}%", f"%{search}%")
        )
    else:
        cur.execute("SELECT * FROM inst_afp ORDER BY nombre_afp")
    rows = cur.fetchall(); conn.close(); return rows

def afp_get(id_afp):
    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()
    cur.execute("SELECT * FROM inst_afp WHERE id_afp = ?", (id_afp,))
    row = cur.fetchone(); conn.close(); return row

def afp_get_by_cod_prev(cod_prev, exclude_id=None):
    """Retorna la AFP que ya tiene ese codPrev_afp, excluyendo el id actual (para edicion)."""
    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()
    if exclude_id:
        cur.execute("SELECT * FROM inst_afp WHERE codPrev_afp = ? AND id_afp != ?", (cod_prev, exclude_id))
    else:
        cur.execute("SELECT * FROM inst_afp WHERE codPrev_afp = ?", (cod_prev,))
    row = cur.fetchone(); conn.close(); return row

def afp_insert(values):
    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()
    col_names = [c[0] for c in AFP_COLUMNS]
    ph = ", ".join("?" * len(col_names))
    cur.execute(f"INSERT INTO inst_afp ({', '.join(col_names)}) VALUES ({ph})", values)
    conn.commit(); conn.close()

def afp_update(id_afp, values):
    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()
    col_names = [c[0] for c in AFP_COLUMNS if c[0] != "id_afp"]
    set_clause = ", ".join(f"{c} = ?" for c in col_names)
    cur.execute(f"UPDATE inst_afp SET {set_clause} WHERE id_afp = ?", values + [id_afp])
    conn.commit(); conn.close()

def afp_delete(id_afp):
    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()
    cur.execute("DELETE FROM inst_afp WHERE id_afp = ?", (id_afp,))
    conn.commit(); conn.close()

def previred_get_all():
    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()
    cur.execute("SELECT codPrev_afp, nombre_afp FROM afp_previred ORDER BY codPrev_afp")
    rows = cur.fetchall(); conn.close(); return rows


# ─────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────

GRP_BLUE = (
    "QGroupBox{font-weight:500;color:#185FA5;border:1px solid #e2e8f0;"
    "border-radius:8px;margin-top:8px;padding-top:8px;background:white;}"
    "QGroupBox::title{subcontrol-origin:margin;left:10px;"
    "padding:0 6px;color:#185FA5;font-size:13px;}"
)
GRP_GREEN = (
    "QGroupBox{font-weight:bold;color:#065f46;border:1px solid #6ee7b7;"
    "border-radius:6px;margin-top:8px;padding-top:8px;}"
    "QGroupBox::title{subcontrol-origin:margin;left:10px;}"
)
STYLE = """
QMainWindow, QWidget { font-family:'Segoe UI',Arial,sans-serif; font-size:13px; background:#f8f9fb; }
QTabWidget::pane { border:1px solid #e2e8f0; border-radius:0 8px 8px 8px; background:white; }
QTabBar::tab {
    background:#f1f5f9; color:#64748b; font-weight:normal;
    padding:8px 20px; border-top-left-radius:6px; border-top-right-radius:6px;
    margin-right:2px; font-size:13px;
    border:1px solid #e2e8f0; border-bottom:none;
}
QTabBar::tab:selected { background:#185FA5; color:white; font-weight:500; border-color:#185FA5; }
QTabBar::tab:hover:!selected { background:#e2e8f0; color:#1e293b; }
QTableWidget {
    background:white; alternate-background-color:#f8fafc;
    gridline-color:#e2e8f0; border:1px solid #e2e8f0; border-radius:0;
    selection-background-color:#EBF4FF; selection-color:#185FA5;
}
QTableWidget::item { padding:6px 12px; }
QTableWidget::item:selected { background:#EBF4FF; color:#185FA5; font-weight:500; }
QHeaderView::section {
    background:#f8fafc; color:#64748b; font-weight:500;
    padding:8px 12px; border:none;
    border-bottom:1px solid #e2e8f0; border-right:1px solid #e2e8f0;
    font-size:12px;
}
QPushButton { border-radius:6px; padding:7px 16px; font-size:13px; font-weight:normal; border:1px solid #e2e8f0; }
QPushButton#btn_add    { background:#1D9E75; color:white; border-color:#1D9E75; }
QPushButton#btn_add:hover { background:#178a64; border-color:#178a64; }
QPushButton#btn_edit   { background:white; color:#1e293b; border-color:#e2e8f0; }
QPushButton#btn_edit:hover { background:#f1f5f9; }
QPushButton#btn_delete { background:white; color:#dc2626; border-color:#e2e8f0; }
QPushButton#btn_delete:hover { background:#fef2f2; border-color:#fca5a5; }
QPushButton#btn_refresh { background:white; color:#64748b; border-color:#e2e8f0; }
QPushButton#btn_refresh:hover { background:#f1f5f9; }
QLineEdit#search_box { border:1px solid #e2e8f0; border-radius:6px; padding:6px 12px; background:white; color:#1e293b; }
QLineEdit#search_box:focus { border-color:#185FA5; }
QDoubleSpinBox, QSpinBox, QLineEdit { border:1px solid #cbd5e1; border-radius:5px; padding:4px 8px; background:white; }
QDoubleSpinBox:focus, QSpinBox:focus, QLineEdit:focus { border-color:#185FA5; }
QStatusBar { background:#f8fafc; color:#64748b; font-size:12px; border-top:1px solid #e2e8f0; }
QPushButton#btn_salir { background:#dc2626; color:white; border-color:#dc2626; }
QPushButton#btn_salir:hover { background:#b91c1c; border-color:#b91c1c; }
QPushButton#btn_exportar { background:#1D9E75; color:white; border-color:#1D9E75; }
QPushButton#btn_exportar:hover { background:#178a64; }
QPushButton#btn_pdf { background:white; color:#185FA5; border-color:#185FA5; }
QPushButton#btn_pdf:hover { background:#EBF4FF; }
QGroupBox { font-weight:500; color:#1e293b; border:1px solid #e2e8f0; border-radius:8px; margin-top:8px; padding-top:8px; }
QGroupBox::title { subcontrol-origin:margin; left:10px; padding:0 6px; color:#185FA5; }
QComboBox { border:1px solid #cbd5e1; border-radius:5px; padding:4px 8px; background:white; }
QComboBox:focus { border-color:#185FA5; }
QScrollBar:vertical { background:#f8fafc; width:8px; border-radius:4px; }
QScrollBar::handle:vertical { background:#cbd5e1; border-radius:4px; min-height:20px; }
QScrollBar::handle:vertical:hover { background:#94a3b8; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height:0; }
"""


# ─────────────────────────────────────────────
# HELPERS UI
# ─────────────────────────────────────────────

def make_decimal(decimals=2, max_val=9_999_999.0):
    w = QDoubleSpinBox()
    w.setRange(0, max_val); w.setDecimals(decimals)
    w.setSingleStep(0.01); w.setGroupSeparatorShown(True)
    w.setMinimumHeight(32); return w

def make_int(max_val=99_999_999):
    w = QSpinBox()
    w.setRange(0, max_val); w.setGroupSeparatorShown(True)
    w.setMinimumHeight(32); return w

def readonly_line(placeholder=""):
    w = QLineEdit()
    w.setPlaceholderText(placeholder); w.setReadOnly(True)
    w.setMinimumHeight(32)
    w.setStyleSheet(
        "background:#eef2f7;color:#374151;border:1px solid #cbd5e1;"
        "border-radius:5px;padding:4px 8px;font-weight:bold;"
    )
    return w

def make_line(placeholder="", max_len=100):
    w = QLineEdit()
    w.setPlaceholderText(placeholder); w.setMaxLength(max_len)
    w.setMinimumHeight(32); return w


# ─────────────────────────────────────────────
# DIÁLOGO PARÁMETROS MENSUALES
# ─────────────────────────────────────────────

class ParamDialog(QDialog):
    def __init__(self, parent=None, record=None):
        super().__init__(parent)
        self.is_edit = record is not None
        self.record  = record
        self.setWindowTitle("✏️ Editar Parámetro" if self.is_edit else "➕ Nuevo Parámetro")
        self.setMinimumWidth(680); self.setModal(True)
        self._build_ui()
        if self.is_edit:
            self._fill_data()
        else:
            last = param_get_last_mes()
            if last:
                self.f_mes_proc.setText(siguiente_mes(last))
        for w in [self.f_uf_mes, self.f_tope_imp_uf, self.f_tope_ces_uf,
                  self.f_sis, self.f_aporte_ccaf]:
            w.valueChanged.connect(self._recalc)
        self.f_mes_proc.textChanged.connect(self._recalc)
        self._recalc()

    def _build_ui(self):
        main = QVBoxLayout(self)
        main.setSpacing(12); main.setContentsMargins(20, 16, 20, 12)

        title = QLabel("✏️ Editar Parámetro" if self.is_edit else "➕ Nuevo Parámetro")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title.setStyleSheet("color:#1e3a5f;"); main.addWidget(title)

        grp_in = QGroupBox("📝  Datos ingresados por el usuario")
        grp_in.setStyleSheet(GRP_BLUE)
        grid_in = QGridLayout(grp_in)
        grid_in.setSpacing(8)
        grid_in.setColumnMinimumWidth(0, 190); grid_in.setColumnMinimumWidth(2, 190)

        self.f_mes_proc    = make_line("YYYY-MM", 7)
        self.f_uf_mes      = make_decimal(2)
        self.f_tope_imp_uf = make_decimal(2)
        self.f_tope_ces_uf = make_decimal(2)
        self.f_sis         = make_decimal(2)
        self.f_imm         = make_int()
        self.f_tope_gratif = make_int()
        self.f_monto_utm   = make_int()
        self.f_aporte_ccaf = make_decimal(2, 7.0)
        self.f_fecha       = make_line("dd-mm-aaaa", 10)
        self.f_aporte_afp  = make_decimal(1, 100.0)
        self.f_seg_social  = make_decimal(1, 100.0)

        if self.is_edit:
            self.f_mes_proc.setReadOnly(True)
            self.f_mes_proc.setStyleSheet(
                "background:#eef2f7;color:#374151;border:1px solid #cbd5e1;border-radius:5px;padding:4px 8px;"
            )

        for lbl_text, widget, row, col in [
            ("Mes Proceso (YYYY-MM) *",      self.f_mes_proc,    0, 0),
            ("UF del Mes *",                 self.f_uf_mes,      0, 2),
            ("Tope Imponible UF AFP *",      self.f_tope_imp_uf, 1, 0),
            ("Tope Cesantía UF *",           self.f_tope_ces_uf, 1, 2),
            ("SIS (%) *",                    self.f_sis,         2, 0),
            ("IMM *",                        self.f_imm,         2, 2),
            ("Tope Gratificación *",         self.f_tope_gratif, 3, 0),
            ("Monto UTM *",                  self.f_monto_utm,   3, 2),
            ("Aporte CCAF (%) *",            self.f_aporte_ccaf, 4, 0),
            ("Formato Fecha (dd-mm-aaaa) *", self.f_fecha,       4, 2),
            ("Aporte AFP (%) *",             self.f_aporte_afp,  5, 0),
            ("Seg. Social Exp. Vida *",      self.f_seg_social,  5, 2),
        ]:
            lbl = QLabel(lbl_text)
            lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            grid_in.addWidget(lbl, row, col); grid_in.addWidget(widget, row, col + 1)
        main.addWidget(grp_in)

        grp_calc = QGroupBox("🔢  Campos calculados automáticamente")
        grp_calc.setStyleSheet(GRP_GREEN)
        grid_calc = QGridLayout(grp_calc)
        grid_calc.setSpacing(8)
        grid_calc.setColumnMinimumWidth(0, 190); grid_calc.setColumnMinimumWidth(2, 190)

        self.c_tope_imp_pesos = readonly_line("ROUND(UF × Tope Imp. UF AFP)")
        self.c_tope_ces_pesos = readonly_line("ROUND(UF × Tope Ces. UF)")
        self.c_factor_sis     = readonly_line("SIS ÷ 100")
        self.c_tope_salud_uf  = readonly_line("Tope Imp. UF × 0.07")
        self.c_tope_salud_p   = readonly_line("ROUND(Tope Salud UF × UF)")
        self.c_ult_dia        = readonly_line("Según mes")
        self.c_aporte_fonasa  = readonly_line("7 − Aporte CCAF")

        for lbl_text, widget, row, col in [
            ("Tope Imponible $ AFP", self.c_tope_imp_pesos, 0, 0),
            ("Tope Cesantía $",      self.c_tope_ces_pesos, 0, 2),
            ("Factor SIS",           self.c_factor_sis,     1, 0),
            ("Tope Salud UF",        self.c_tope_salud_uf,  1, 2),
            ("Tope Salud $",         self.c_tope_salud_p,   2, 0),
            ("Último Día del Mes",   self.c_ult_dia,         2, 2),
            ("Aporte FONASA (%)",    self.c_aporte_fonasa,  3, 0),
        ]:
            lbl = QLabel(lbl_text)
            lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            grid_calc.addWidget(lbl, row, col); grid_calc.addWidget(widget, row, col + 1)
        main.addWidget(grp_calc)

        nota = QLabel("  * Campos obligatorios   |   Los campos calculados se actualizan en tiempo real")
        nota.setStyleSheet("color:#6b7280;font-size:11px;padding:4px 0;"); main.addWidget(nota)

        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.button(QDialogButtonBox.StandardButton.Ok).setText("💾  Guardar")
        btn_box.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        btn_box.accepted.connect(self._validate_and_accept)
        btn_box.rejected.connect(self.reject)
        main.addWidget(btn_box)

    def _recalc(self):
        try:
            calc = calcular_campos(
                self.f_uf_mes.value(), self.f_tope_imp_uf.value(),
                self.f_tope_ces_uf.value(), self.f_sis.value(),
                self.f_aporte_ccaf.value(), self.f_mes_proc.text().strip()
            )
            self.c_tope_imp_pesos.setText(f"{calc['tope_imp_pesos_afp']:,}")
            self.c_tope_ces_pesos.setText(f"{calc['tope_ces_pesos']:,}")
            self.c_factor_sis.setText(f"{calc['factor_sis']:.4f}")
            self.c_tope_salud_uf.setText(f"{calc['tope_salud_uf']:.3f}")
            self.c_tope_salud_p.setText(f"{calc['tope_salud_pesos']:,}")
            self.c_ult_dia.setText(str(calc['ult_dia_mes']))
            self.c_aporte_fonasa.setText(f"{calc['aporte_fonasa']:.2f}")
        except Exception:
            pass

    def _fill_data(self):
        idx = {c[0]: i for i, c in enumerate(PARAM_COLUMNS)}
        r   = self.record
        self.f_mes_proc.setText(str(r[idx["mes_proc"]] or ""))
        self.f_uf_mes.setValue(float(r[idx["uf_mes"]] or 0))
        self.f_tope_imp_uf.setValue(float(r[idx["tope_imp_uf_afp"]] or 0))
        self.f_tope_ces_uf.setValue(float(r[idx["tope_ces_uf"]] or 0))
        self.f_sis.setValue(float(r[idx["sis"]] or 0))
        self.f_imm.setValue(int(r[idx["imm"]] or 0))
        self.f_tope_gratif.setValue(int(r[idx["tope_gratif"]] or 0))
        self.f_monto_utm.setValue(int(r[idx["monto_utm"]] or 0))
        self.f_aporte_ccaf.setValue(float(r[idx["aporte_ccaf"]] or 0))
        self.f_fecha.setText(str(r[idx["formato_fecha"]] or ""))
        self.f_aporte_afp.setValue(float(r[idx["aporte_afp"]] or 0))
        self.f_seg_social.setValue(float(r[idx["seg_social_exp_vida"]] or 0))

    def _validate_and_accept(self):
        errors = []
        mes = self.f_mes_proc.text().strip()
        if not mes or len(mes) != 7 or mes[4] != "-":
            errors.append("• Mes Proceso: formato inválido — use YYYY-MM (ej: 2025-04)")
        else:
            try:
                y, m = int(mes[:4]), int(mes[5:])
                if not (1 <= m <= 12) or y < 2000:
                    raise ValueError
            except ValueError:
                errors.append("• Mes Proceso: año o mes inválido")

        for widget, name in [
            (self.f_uf_mes, "UF del Mes"), (self.f_tope_imp_uf, "Tope Imponible UF AFP"),
            (self.f_tope_ces_uf, "Tope Cesantía UF"), (self.f_sis, "SIS"),
            (self.f_aporte_ccaf, "Aporte CCAF"), (self.f_aporte_afp, "Aporte AFP"),
            (self.f_seg_social, "Seg. Social Exp. Vida"),
        ]:
            if widget.value() == 0:
                errors.append(f"• {name}: no puede ser cero ni estar vacío")

        for widget, name in [
            (self.f_imm, "IMM"), (self.f_tope_gratif, "Tope Gratificación"), (self.f_monto_utm, "Monto UTM")
        ]:
            if widget.value() == 0:
                errors.append(f"• {name}: no puede ser cero ni estar vacío")

        if self.f_aporte_ccaf.value() > 7:
            errors.append("• Aporte CCAF: no puede superar 7%")

        fecha = self.f_fecha.text().strip()
        if not fecha:
            errors.append("• Formato Fecha: campo obligatorio")
        else:
            try:
                datetime.strptime(fecha, "%d-%m-%Y")
            except ValueError:
                errors.append("• Formato Fecha: use dd-mm-aaaa (ej: 31-03-2025)")

        if errors:
            QMessageBox.warning(self, "⚠️ Errores de validación",
                                "Por favor corrija los siguientes errores:\n\n" + "\n".join(errors))
            return
        self.accept()

    def get_values(self):
        mes, uf = self.f_mes_proc.text().strip(), self.f_uf_mes.value()
        tiuf, tcuf = self.f_tope_imp_uf.value(), self.f_tope_ces_uf.value()
        sis, ccaf  = self.f_sis.value(), self.f_aporte_ccaf.value()
        calc = calcular_campos(uf, tiuf, tcuf, sis, ccaf, mes)
        return [
            mes, uf, tiuf, calc["tope_imp_pesos_afp"],
            tcuf, calc["tope_ces_pesos"], sis, calc["factor_sis"],
            calc["tope_salud_uf"], calc["tope_salud_pesos"],
            self.f_imm.value(), self.f_tope_gratif.value(), self.f_monto_utm.value(),
            calc["ult_dia_mes"], ccaf, calc["aporte_fonasa"],
            self.f_fecha.text().strip(), self.f_aporte_afp.value(), self.f_seg_social.value(),
        ]


# ─────────────────────────────────────────────
# DIÁLOGO AFP
# ─────────────────────────────────────────────

class AfpDialog(QDialog):
    def __init__(self, parent=None, record=None):
        super().__init__(parent)
        self.is_edit = record is not None
        self.record  = record
        # Cargar tabla Previred como dict {cod: nombre}
        self.previred = {cod: nombre for cod, nombre in previred_get_all()}
        self.setWindowTitle("✏️ Editar AFP" if self.is_edit else "➕ Nueva AFP")
        self.setMinimumWidth(580); self.setModal(True)
        self._build_ui()
        if self.is_edit:
            self._fill_data()

    def _build_ui(self):
        main = QVBoxLayout(self)
        main.setSpacing(12); main.setContentsMargins(20, 16, 20, 12)

        title = QLabel("✏️ Editar AFP" if self.is_edit else "➕ Nueva AFP")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title.setStyleSheet("color:#1e3a5f;"); main.addWidget(title)

        grp = QGroupBox("📝  Datos de la AFP")
        grp.setStyleSheet(GRP_BLUE)
        grid = QGridLayout(grp)
        grid.setSpacing(10)
        grid.setColumnMinimumWidth(0, 160)
        grid.setColumnMinimumWidth(1, 320)

        self.f_id     = make_line("ej: capital", 30)
        self.f_nombre = make_line("ej: AFP Capital", 80)
        self.f_rut    = make_line("ej: 98000000-1", 20)
        self.f_cot    = make_decimal(2, 100.0)
        self.f_observ = QTextEdit()
        self.f_observ.setMaximumHeight(60)

        if self.is_edit:
            self.f_id.setReadOnly(True)
            self.f_id.setStyleSheet(
                "background:#eef2f7;color:#374151;border:1px solid #cbd5e1;border-radius:5px;padding:4px 8px;"
            )

        # Campo código Previred con indicador visual en tiempo real
        cod_container = QWidget()
        cod_layout    = QHBoxLayout(cod_container)
        cod_layout.setContentsMargins(0, 0, 0, 0); cod_layout.setSpacing(8)
        self.f_cod_prev = make_line("ej: 33", 10)
        self.f_cod_prev.setMaximumWidth(90)
        self.lbl_prev_status = QLabel("")
        self.lbl_prev_status.setMinimumWidth(220)
        cod_layout.addWidget(self.f_cod_prev)
        cod_layout.addWidget(self.lbl_prev_status)
        cod_layout.addStretch()
        self.f_cod_prev.textChanged.connect(self._check_previred)

        for lbl_text, widget, row in [
            ("ID AFP *",            self.f_id,       0),
            ("Nombre AFP *",        self.f_nombre,   2),
            ("RUT AFP",             self.f_rut,      3),
            ("Cód. Previred *",     cod_container,   4),
            ("Cotización (%) *",    self.f_cot,      5),
            ("Observaciones",       self.f_observ,   6),
        ]:
            lbl = QLabel(lbl_text)
            lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            grid.addWidget(lbl, row, 0); grid.addWidget(widget, row, 1)

        # Referencia de códigos válidos
        prev_list = "  |  ".join(
            f"{cod} = {nombre}"
            for cod, nombre in sorted(self.previred.items(), key=lambda x: x[0])
        )
        lbl_ref = QLabel(f"Códigos Previred válidos:\n{prev_list}")
        lbl_ref.setWordWrap(True)
        lbl_ref.setStyleSheet(
            "color:#1d4ed8; font-size:11px; background:#eff6ff;"
            "border:1px solid #bfdbfe; border-radius:5px; padding:6px 8px;"
        )
        grid.addWidget(lbl_ref, 7, 0, 1, 2)
        main.addWidget(grp)

        nota = QLabel("  * Campos obligatorios   |   El código Previred se normaliza a 2 dígitos automáticamente")
        nota.setStyleSheet("color:#6b7280;font-size:11px;"); main.addWidget(nota)

        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.button(QDialogButtonBox.StandardButton.Ok).setText("💾  Guardar")
        btn_box.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        btn_box.accepted.connect(self._validate_and_accept)
        btn_box.rejected.connect(self.reject)
        main.addWidget(btn_box)

    def _check_previred(self, text):
        """Valida el código Previred en tiempo real, normalizando a 2 chars."""
        cod = text.strip()
        if not cod:
            self.lbl_prev_status.setText("")
            return
        cod_norm = normalizar_cod(cod)
        nombre   = self.previred.get(cod_norm)
        if nombre:
            self.lbl_prev_status.setText(f"✅  {cod_norm}  →  {nombre}")
            self.lbl_prev_status.setStyleSheet("color:#16a34a; font-weight:bold;")
        else:
            self.lbl_prev_status.setText(f"❌  '{cod_norm}' no existe en Previred")
            self.lbl_prev_status.setStyleSheet("color:#dc2626; font-weight:bold;")

    def _fill_data(self):
        idx = {c[0]: i for i, c in enumerate(AFP_COLUMNS)}
        r   = self.record
        self.f_id.setText(str(r[idx["id_afp"]] or ""))
        self.f_nombre.setText(str(r[idx["nombre_afp"]] or ""))
        self.f_rut.setText(str(r[idx["rut_afp"]] or ""))
        self.f_cod_prev.setText(str(r[idx["codPrev_afp"]] or ""))
        self.f_cot.setValue(float(r[idx["cot_afp"]] or 0))
        self.f_observ.setPlainText(str(r[idx["observ_afp"]] or ""))

    def _validate_and_accept(self):
        errors = []
        if not self.f_id.text().strip():
            errors.append("• ID AFP: campo obligatorio")
        if not self.f_nombre.text().strip():
            errors.append("• Nombre AFP: campo obligatorio")
        if self.f_cot.value() == 0:
            errors.append("• Cotización: no puede ser cero")

        cod = self.f_cod_prev.text().strip()
        if not cod:
            errors.append("• Código Previred: campo obligatorio")
        else:
            cod_norm = normalizar_cod(cod)
            if cod_norm not in self.previred:
                errors.append(
                    f"• Código Previred '{cod_norm}': no existe en la tabla Previred.\n"
                    f"  Códigos válidos: {', '.join(sorted(self.previred.keys()))}"
                )

        if errors:
            QMessageBox.warning(self, "⚠️ Errores de validación",
                                "Por favor corrija los siguientes errores:\n\n" + "\n".join(errors))
            return
        self.accept()

    def get_values(self):
        cod_norm = normalizar_cod(self.f_cod_prev.text().strip())
        return [
            self.f_id.text().strip(),
            "af",
            self.f_nombre.text().strip(),
            self.f_rut.text().strip(),
            cod_norm,           # siempre normalizado con 2 dígitos mínimo
            self.f_cot.value(),
            self.f_observ.toPlainText().strip(),
        ]


# ─────────────────────────────────────────────
# WIDGET TABLA GENÉRICO
# ─────────────────────────────────────────────

class TableWidget(QWidget):
    def __init__(self, headers, search_placeholder="Buscar..."):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 8); layout.setSpacing(8)

        bar = QHBoxLayout()
        self.btn_add     = QPushButton("➕  Nuevo");    self.btn_add.setObjectName("btn_add");     self.btn_add.setFixedHeight(36)
        self.btn_edit    = QPushButton("✏️  Editar");   self.btn_edit.setObjectName("btn_edit");    self.btn_edit.setFixedHeight(36)
        self.btn_delete  = QPushButton("🗑️  Eliminar"); self.btn_delete.setObjectName("btn_delete"); self.btn_delete.setFixedHeight(36)
        self.btn_refresh = QPushButton("🔄  Actualizar"); self.btn_refresh.setObjectName("btn_refresh"); self.btn_refresh.setFixedHeight(36)

        bar.addWidget(self.btn_add); bar.addWidget(self.btn_edit)
        bar.addWidget(self.btn_delete); bar.addStretch()
        bar.addWidget(QLabel("🔍 Buscar:"))
        self.search_box = QLineEdit()
        self.search_box.setObjectName("search_box")
        self.search_box.setPlaceholderText(search_placeholder)
        self.search_box.setFixedWidth(200)
        bar.addWidget(self.search_box); bar.addWidget(self.btn_refresh)
        layout.addLayout(bar)

        self.table = QTableWidget()
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)

    def selected_key(self, col=0):
        row = self.table.currentRow()
        if row < 0 or not self.table.item(row, col):
            return None
        return self.table.item(row, col).text()


# ─────────────────────────────────────────────
# VENTANA PRINCIPAL
# ─────────────────────────────────────────────



# ─────────────────────────────────────────────

# DIÁLOGO INSTITUCIONES SALUD

# ─────────────────────────────────────────────



# ─────────────────────────────────────────────
# DIALOGO INSTITUCIONES APV
# ─────────────────────────────────────────────

# ─────────────────────────────────────────────
# DIALOGO INSTITUCIONES CAJAS
# ─────────────────────────────────────────────

# ─────────────────────────────────────────────

# ─────────────────────────────────────────────
# VALIDACION RUT Y DIALOGO EMPRESAS
# ─────────────────────────────────────────────

def validar_rut(rut):
    rut=rut.upper().replace(".","").replace("-","").strip()
    if len(rut)<2: return False
    cuerpo,dv=rut[:-1],rut[-1]
    if not cuerpo.isdigit(): return False
    suma,mult=0,2
    for c in reversed(cuerpo):
        suma+=int(c)*mult
        mult=mult+1 if mult<7 else 2
    resto=suma%11
    dv_calc=str(11-resto) if resto not in (0,1) else ("0" if resto==0 else "K")
    return dv==dv_calc

def formatear_rut(rut):
    rut=rut.upper().replace(".","").replace("-","").strip()
    if len(rut)<2: return rut
    cuerpo,dv=rut[:-1],rut[-1]
    if not cuerpo.isdigit(): return rut
    return f"{int(cuerpo):,}".replace(",",".")+"-"+dv

class EmpresaDialog(QDialog):
    def __init__(self, parent=None, record=None):
        super().__init__(parent)
        self.is_edit = record is not None
        self.record  = record
        self.setWindowTitle("Editar Empresa" if self.is_edit else "Nueva Empresa")
        self.setMinimumWidth(600); self.setMinimumHeight(440); self.setModal(True)
        self._build_ui()
        if self.is_edit:
            self._fill_data()

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setSpacing(8); outer.setContentsMargins(14,10,14,10)
        title = QLabel("Editar Empresa" if self.is_edit else "Nueva Empresa")
        title.setFont(QFont("Segoe UI",13,QFont.Weight.Bold))
        title.setStyleSheet("color:#1e3a5f;")
        outer.addWidget(title)
        self.tabs = QTabWidget()
        outer.addWidget(self.tabs)

        # Pestana 1: Identificacion
        tab1=QWidget(); g1=QGridLayout(tab1)
        g1.setSpacing(10); g1.setContentsMargins(14,14,14,14)
        g1.setColumnMinimumWidth(0,150); g1.setColumnMinimumWidth(1,350)
        self.f_id=make_line("ej: emp001",30)
        self.f_rut=make_line("ej: 76354771-9",15)
        self.lbl_rut=QLabel(""); self.lbl_rut.setMinimumWidth(160)
        self.f_razon=make_line("ej: Mi Empresa S.A.",100)
        self.f_fant=make_line("ej: Mi Empresa",100)
        self.f_cod=make_line("ej: 620100",10)
        if self.is_edit:
            self.f_id.setReadOnly(True)
            self.f_id.setStyleSheet("background:#eef2f7;color:#374151;border:1px solid #cbd5e1;border-radius:5px;padding:4px 8px;")
        rut_c=QWidget(); rl=QHBoxLayout(rut_c)
        rl.setContentsMargins(0,0,0,0); rl.setSpacing(6)
        rl.addWidget(self.f_rut); rl.addWidget(self.lbl_rut); rl.addStretch()
        self.f_rut.textChanged.connect(self._check_rut)
        for lbl,w,row in [("ID Empresa *",self.f_id,0),("RUT Empresa *",rut_c,1),
                           ("Razon Social *",self.f_razon,2),("Nombre Fantasia",self.f_fant,3),
                           ("Cod. Act. Economica",self.f_cod,4)]:
            lb=QLabel(lbl); lb.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
            g1.addWidget(lb,row,0); g1.addWidget(w,row,1)
        self.tabs.addTab(tab1,"Identificacion")

        # Pestana 2: Ubicacion
        tab2=QWidget(); g2=QGridLayout(tab2)
        g2.setSpacing(10); g2.setContentsMargins(14,14,14,14)
        g2.setColumnMinimumWidth(0,150); g2.setColumnMinimumWidth(1,350)
        self.cb_region=QComboBox(); self.cb_comuna=QComboBox(); self.cb_ciudad=QComboBox()
        self.f_direcc=make_line("ej: Av. Principal 123",100)
        self.regiones_data=regiones_get_all()
        self.cb_region.addItem("-- Seleccione Region --","")
        for num,nom in self.regiones_data: self.cb_region.addItem(nom,num)
        self.cb_region.currentIndexChanged.connect(self._on_region_changed)
        self.cb_comuna.currentIndexChanged.connect(self._on_comuna_changed)
        for lbl,w,row in [("Region *",self.cb_region,0),("Comuna *",self.cb_comuna,1),
                           ("Ciudad *",self.cb_ciudad,2),("Direccion",self.f_direcc,3)]:
            lb=QLabel(lbl); lb.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
            g2.addWidget(lb,row,0); g2.addWidget(w,row,1)
        self.tabs.addTab(tab2,"Ubicacion")

        # Pestana 3: Prevision
        tab3=QWidget(); g3=QGridLayout(tab3)
        g3.setSpacing(10); g3.setContentsMargins(14,14,14,14)
        g3.setColumnMinimumWidth(0,150); g3.setColumnMinimumWidth(1,350)
        self.cb_ccaf=QComboBox(); self.cb_mutual=QComboBox()
        self.f_cot_mut=make_line("ej: 0.93",10)
        conn=sqlite3.connect(DB_PATH); cur=conn.cursor()
        self.cb_ccaf.addItem("-- Sin CCAF --","")
        for _,n in cur.execute("SELECT id_inst,nombre_inst FROM inst_cajas ORDER BY nombre_inst"): self.cb_ccaf.addItem(n,n)
        self.cb_mutual.addItem("-- Seleccione Mutual --","")
        for _,n in cur.execute("SELECT id_inst,nombre_inst FROM inst_mutuales ORDER BY nombre_inst"): self.cb_mutual.addItem(n,n)
        conn.close()
        for lbl,w,row in [("CCAF",self.cb_ccaf,0),("Mutual *",self.cb_mutual,1),("Cot. Mutual %",self.f_cot_mut,2)]:
            lb=QLabel(lbl); lb.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
            g3.addWidget(lb,row,0); g3.addWidget(w,row,1)
        self.tabs.addTab(tab3,"Prevision")

        # Pestana 4: Rep. Legal
        tab4=QWidget(); g4=QGridLayout(tab4)
        g4.setSpacing(10); g4.setContentsMargins(14,14,14,14)
        g4.setColumnMinimumWidth(0,150); g4.setColumnMinimumWidth(1,350)
        self.f_rut_rep=make_line("ej: 12345678-9",15)
        self.lbl_rut_rep=QLabel(""); self.lbl_rut_rep.setMinimumWidth(160)
        self.f_nom_rep=make_line("ej: Juan Perez",80)
        self.f_email_rep=make_line("ej: juan@empresa.cl",100)
        self.f_email_emp=make_line("ej: contacto@empresa.cl",100)
        self.f_fono=make_line("ej: +56912345678",20)
        self.f_rut_rep.textChanged.connect(self._check_rut_rep)
        rut_rep_c=QWidget(); rrl=QHBoxLayout(rut_rep_c)
        rrl.setContentsMargins(0,0,0,0); rrl.setSpacing(6)
        rrl.addWidget(self.f_rut_rep); rrl.addWidget(self.lbl_rut_rep); rrl.addStretch()
        for lbl,w,row in [("RUT Rep. Legal *",rut_rep_c,0),("Nombre Rep. Legal *",self.f_nom_rep,1),
                           ("Email Rep. Legal",self.f_email_rep,2),("Email Empresa",self.f_email_emp,3),
                           ("Telefono",self.f_fono,4)]:
            lb=QLabel(lbl); lb.setAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
            g4.addWidget(lb,row,0); g4.addWidget(w,row,1)
        self.tabs.addTab(tab4,"Rep. Legal")

        nota=QLabel("  * Campos obligatorios"); nota.setStyleSheet("color:#6b7280;font-size:11px;")
        outer.addWidget(nota)
        btn_box=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        btn_box.button(QDialogButtonBox.StandardButton.Ok).setText("Guardar")
        btn_box.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        btn_box.accepted.connect(self._validate_and_accept); btn_box.rejected.connect(self.reject)
        outer.addWidget(btn_box)

    def _check_rut(self,text):
        if not text.strip(): self.lbl_rut.setText(""); return
        if validar_rut(text.strip()):
            self.lbl_rut.setText(f"OK  {formatear_rut(text.strip())}"); self.lbl_rut.setStyleSheet("color:#16a34a;font-weight:bold;")
        else:
            self.lbl_rut.setText("RUT invalido"); self.lbl_rut.setStyleSheet("color:#dc2626;font-weight:bold;")

    def _check_rut_rep(self,text):
        if not text.strip(): self.lbl_rut_rep.setText(""); return
        if validar_rut(text.strip()):
            self.lbl_rut_rep.setText(f"OK  {formatear_rut(text.strip())}"); self.lbl_rut_rep.setStyleSheet("color:#16a34a;font-weight:bold;")
        else:
            self.lbl_rut_rep.setText("RUT invalido"); self.lbl_rut_rep.setStyleSheet("color:#dc2626;font-weight:bold;")

    def _on_region_changed(self,idx):
        num=self.cb_region.currentData()
        self.cb_comuna.clear(); self.cb_ciudad.clear()
        if not num: return
        self.cb_comuna.addItem("-- Seleccione Comuna --","")
        for c in comunas_get_by_region(num): self.cb_comuna.addItem(c,c)

    def _on_comuna_changed(self,idx):
        num=self.cb_region.currentData(); comuna=self.cb_comuna.currentData()
        self.cb_ciudad.clear()
        if not num or not comuna: return
        self.cb_ciudad.addItem("-- Seleccione Ciudad --","")
        for c in ciudades_get_by_comuna(num,comuna): self.cb_ciudad.addItem(c,c)

    def _fill_data(self):
        r=self.record
        self.f_id.setText(str(r[0] or "")); self.f_rut.setText(str(r[1] or ""))
        self.f_razon.setText(str(r[2] or "")); self.f_fant.setText(str(r[3] or ""))
        for i in range(self.cb_region.count()):
            if self.cb_region.itemText(i)==str(r[4] or ""): self.cb_region.setCurrentIndex(i); break
        QApplication.processEvents()
        for i in range(self.cb_comuna.count()):
            if self.cb_comuna.itemText(i)==str(r[5] or ""): self.cb_comuna.setCurrentIndex(i); break
        QApplication.processEvents()
        for i in range(self.cb_ciudad.count()):
            if self.cb_ciudad.itemText(i)==str(r[6] or ""): self.cb_ciudad.setCurrentIndex(i); break
        self.f_direcc.setText(str(r[7] or "")); self.f_fono.setText(str(r[8] or ""))
        self.f_email_rep.setText(str(r[9] or ""))
        for cb,val in [(self.cb_ccaf,r[10]),(self.cb_mutual,r[11])]:
            for i in range(cb.count()):
                if cb.itemData(i)==str(val or ""): cb.setCurrentIndex(i); break
        self.f_cot_mut.setText(str(r[12] or ""))
        self.f_cod.setText(str(r[13] or ""))
        self.f_rut_rep.setText(str(r[14] or ""))
        self.f_nom_rep.setText(str(r[15] or ""))
        self.f_email_emp.setText(str(r[16] or ""))

    def _validate_and_accept(self):
        errors=[]
        if not self.f_id.text().strip(): errors.append("- ID Empresa: obligatorio")
        if not self.f_rut.text().strip(): errors.append("- RUT Empresa: obligatorio")
        elif not validar_rut(self.f_rut.text().strip()): errors.append("- RUT Empresa: invalido")
        if not self.f_razon.text().strip(): errors.append("- Razon Social: obligatoria")
        if not self.cb_region.currentData(): errors.append("- Region: obligatoria (pestana Ubicacion)")
        if not self.cb_comuna.currentData(): errors.append("- Comuna: obligatoria (pestana Ubicacion)")
        if not self.cb_ciudad.currentData(): errors.append("- Ciudad: obligatoria (pestana Ubicacion)")
        if not self.cb_mutual.currentData(): errors.append("- Mutual: obligatoria (pestana Prevision)")
        if not self.f_rut_rep.text().strip(): errors.append("- RUT Rep. Legal: obligatorio")
        elif not validar_rut(self.f_rut_rep.text().strip()): errors.append("- RUT Rep. Legal: invalido")
        if not self.f_nom_rep.text().strip(): errors.append("- Nombre Rep. Legal: obligatorio")
        if errors:
            QMessageBox.warning(self,"Errores","Corrija los siguientes errores:\n\n"+"\n".join(errors)); return
        self.accept()

    def get_values(self):
        return [self.f_id.text().strip(),self.f_rut.text().strip(),self.f_razon.text().strip(),
                self.f_fant.text().strip(),self.cb_region.currentText(),self.cb_comuna.currentText(),
                self.cb_ciudad.currentText(),self.f_direcc.text().strip(),self.f_fono.text().strip(),
                self.f_email_rep.text().strip(),self.cb_ccaf.currentData() or "",
                self.cb_mutual.currentData() or "",self.f_cot_mut.text().strip(),
                self.f_cod.text().strip(),self.f_rut_rep.text().strip(),
                self.f_nom_rep.text().strip(),self.f_email_emp.text().strip()]


class CajasDialog(QDialog):
    def __init__(self, parent=None, record=None):
        super().__init__(parent)
        self.is_edit = record is not None
        self.record  = record
        self.setWindowTitle("Editar CCAF" if self.is_edit else "Nueva CCAF")
        self.setMinimumWidth(580); self.setModal(True)
        self._build_ui()
        if self.is_edit:
            self._fill_data()

    def _build_ui(self):
        main = QVBoxLayout(self)
        main.setSpacing(12); main.setContentsMargins(20, 16, 20, 12)
        title = QLabel("Editar CCAF" if self.is_edit else "Nueva CCAF")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title.setStyleSheet("color:#1e3a5f;"); main.addWidget(title)
        grp = QGroupBox("Datos de la Caja de Compensacion")
        grp.setStyleSheet(GRP_BLUE)
        grid = QGridLayout(grp)
        grid.setSpacing(10)
        grid.setColumnMinimumWidth(0, 160); grid.setColumnMinimumWidth(1, 320)
        self.f_id    = make_line("ej: losandes", 30)
        self.f_clas  = make_line("ej: ca", 10)
        self.f_nom   = make_line("ej: Caja Los Andes", 80)
        self.f_doc   = make_line("ej: 81826800-9", 20)
        self.f_cod   = make_line("ej: 01", 10)
        self.f_val   = make_line("ej: 0", 10)
        self.f_val2  = make_line("ej: 0", 10)
        self.f_val3  = make_line("ej: 0", 10)
        self.f_dato  = make_line("", 80)
        if self.is_edit:
            self.f_id.setReadOnly(True)
            self.f_id.setStyleSheet("background:#eef2f7;color:#374151;border:1px solid #cbd5e1;border-radius:5px;padding:4px 8px;")
        for lbl_text, widget, row in [
            ("ID Institucion *", self.f_id,   0),
            ("Clasificacion *",  self.f_clas, 1),
            ("Nombre *",         self.f_nom,  2),
            ("Doc. Identidad",   self.f_doc,  3),
            ("Cod. Equivalente *",self.f_cod, 4),
            ("Valor",            self.f_val,  5),
            ("Valor 2",          self.f_val2, 6),
            ("Valor 3",          self.f_val3, 7),
            ("Dato Adicional",   self.f_dato, 8),
        ]:
            lbl = QLabel(lbl_text)
            lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            grid.addWidget(lbl, row, 0); grid.addWidget(widget, row, 1)
        main.addWidget(grp)
        nota = QLabel("  * Campos obligatorios")
        nota.setStyleSheet("color:#6b7280;font-size:11px;"); main.addWidget(nota)
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.button(QDialogButtonBox.StandardButton.Ok).setText("Guardar")
        btn_box.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        btn_box.accepted.connect(self._validate_and_accept)
        btn_box.rejected.connect(self.reject)
        main.addWidget(btn_box)

    def _fill_data(self):
        r = self.record
        self.f_id.setText(str(r[0] or ""))
        self.f_clas.setText(str(r[1] or ""))
        self.f_nom.setText(str(r[2] or ""))
        self.f_doc.setText(str(r[3] or ""))
        self.f_cod.setText(str(r[4] or ""))
        self.f_val.setText(str(r[5] or ""))
        self.f_val2.setText(str(r[6] or ""))
        self.f_val3.setText(str(r[7] or ""))
        self.f_dato.setText(str(r[8] or ""))

    def _validate_and_accept(self):
        errors = []
        if not self.f_id.text().strip(): errors.append("- ID: campo obligatorio")
        if not self.f_clas.text().strip(): errors.append("- Clasificacion: campo obligatorio")
        if not self.f_nom.text().strip(): errors.append("- Nombre: campo obligatorio")
        if not self.f_cod.text().strip(): errors.append("- Cod. Equivalente: campo obligatorio")
        if errors:
            QMessageBox.warning(self, "Errores", "Corrija los siguientes errores:\n\n" + "\n".join(errors)); return
        self.accept()

    def get_values(self):
        return [self.f_id.text().strip(), self.f_clas.text().strip(), self.f_nom.text().strip(),
                self.f_doc.text().strip(), self.f_cod.text().strip(), self.f_val.text().strip(),
                self.f_val2.text().strip(), self.f_val3.text().strip(), self.f_dato.text().strip()]

# ─────────────────────────────────────────────
# DIALOGO INSTITUCIONES MUTUALES
# ─────────────────────────────────────────────

class MutualesDialog(QDialog):
    def __init__(self, parent=None, record=None):
        super().__init__(parent)
        self.is_edit = record is not None
        self.record  = record
        self.setWindowTitle("Editar Mutual" if self.is_edit else "Nueva Mutual")
        self.setMinimumWidth(580); self.setModal(True)
        self._build_ui()
        if self.is_edit:
            self._fill_data()

    def _build_ui(self):
        main = QVBoxLayout(self)
        main.setSpacing(12); main.setContentsMargins(20, 16, 20, 12)
        title = QLabel("Editar Mutual" if self.is_edit else "Nueva Mutual")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title.setStyleSheet("color:#1e3a5f;"); main.addWidget(title)
        grp = QGroupBox("Datos de la Mutual")
        grp.setStyleSheet(GRP_BLUE)
        grid = QGridLayout(grp)
        grid.setSpacing(10)
        grid.setColumnMinimumWidth(0, 160); grid.setColumnMinimumWidth(1, 320)
        self.f_id    = make_line("ej: achs", 30)
        self.f_clas  = make_line("ej: mu", 10)
        self.f_nom   = make_line("ej: ACHS", 80)
        self.f_doc   = make_line("ej: 70360100-6", 20)
        self.f_cod   = make_line("ej: 01", 10)
        self.f_val   = make_line("ej: 0", 10)
        self.f_val2  = make_line("ej: 0", 10)
        self.f_val3  = make_line("ej: 0", 10)
        self.f_dato  = make_line("", 80)
        if self.is_edit:
            self.f_id.setReadOnly(True)
            self.f_id.setStyleSheet("background:#eef2f7;color:#374151;border:1px solid #cbd5e1;border-radius:5px;padding:4px 8px;")
        for lbl_text, widget, row in [
            ("ID Institucion *",  self.f_id,   0),
            ("Clasificacion *",   self.f_clas, 1),
            ("Nombre *",          self.f_nom,  2),
            ("Doc. Identidad",    self.f_doc,  3),
            ("Cod. Equivalente *",self.f_cod,  4),
            ("Valor",             self.f_val,  5),
            ("Valor 2",           self.f_val2, 6),
            ("Valor 3",           self.f_val3, 7),
            ("Dato Adicional",    self.f_dato, 8),
        ]:
            lbl = QLabel(lbl_text)
            lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            grid.addWidget(lbl, row, 0); grid.addWidget(widget, row, 1)
        main.addWidget(grp)
        nota = QLabel("  * Campos obligatorios")
        nota.setStyleSheet("color:#6b7280;font-size:11px;"); main.addWidget(nota)
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.button(QDialogButtonBox.StandardButton.Ok).setText("Guardar")
        btn_box.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        btn_box.accepted.connect(self._validate_and_accept)
        btn_box.rejected.connect(self.reject)
        main.addWidget(btn_box)

    def _fill_data(self):
        r = self.record
        self.f_id.setText(str(r[0] or ""))
        self.f_clas.setText(str(r[1] or ""))
        self.f_nom.setText(str(r[2] or ""))
        self.f_doc.setText(str(r[3] or ""))
        self.f_cod.setText(str(r[4] or ""))
        self.f_val.setText(str(r[5] or ""))
        self.f_val2.setText(str(r[6] or ""))
        self.f_val3.setText(str(r[7] or ""))
        self.f_dato.setText(str(r[8] or ""))

    def _validate_and_accept(self):
        errors = []
        if not self.f_id.text().strip(): errors.append("- ID: campo obligatorio")
        if not self.f_clas.text().strip(): errors.append("- Clasificacion: campo obligatorio")
        if not self.f_nom.text().strip(): errors.append("- Nombre: campo obligatorio")
        if not self.f_cod.text().strip(): errors.append("- Cod. Equivalente: campo obligatorio")
        if errors:
            QMessageBox.warning(self, "Errores", "Corrija los siguientes errores:\n\n" + "\n".join(errors)); return
        self.accept()

    def get_values(self):
        return [self.f_id.text().strip(), self.f_clas.text().strip(), self.f_nom.text().strip(),
                self.f_doc.text().strip(), self.f_cod.text().strip(), self.f_val.text().strip(),
                self.f_val2.text().strip(), self.f_val3.text().strip(), self.f_dato.text().strip()]

class ApvDialog(QDialog):
    def __init__(self, parent=None, record=None):
        super().__init__(parent)
        self.is_edit = record is not None
        self.record  = record
        self.setWindowTitle("Editar Institucion APV" if self.is_edit else "Nueva Institucion APV")
        self.setMinimumWidth(580); self.setModal(True)
        self._build_ui()
        if self.is_edit:
            self._fill_data()

    def _build_ui(self):
        main = QVBoxLayout(self)
        main.setSpacing(12); main.setContentsMargins(20, 16, 20, 12)
        title = QLabel("Editar Institucion APV" if self.is_edit else "Nueva Institucion APV")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title.setStyleSheet("color:#1e3a5f;"); main.addWidget(title)
        grp = QGroupBox("Datos de la Institucion APV")
        grp.setStyleSheet(GRP_BLUE)
        grid = QGridLayout(grp)
        grid.setSpacing(10)
        grid.setColumnMinimumWidth(0, 160)
        grid.setColumnMinimumWidth(1, 320)
        self.f_id     = make_line("ej: 100", 30)
        self.f_clasif = make_line("ej: ap", 10)
        self.f_nombre = make_line("ej: APV Capital", 80)
        self.f_rut    = make_line("ej: 76354771-9", 20)
        self.f_cod    = make_line("ej: 033", 10)
        if self.is_edit:
            self.f_id.setReadOnly(True)
            self.f_id.setStyleSheet("background:#eef2f7;color:#374151;border:1px solid #cbd5e1;border-radius:5px;padding:4px 8px;")
        for lbl_text, widget, row in [
            ("ID APV *",             self.f_id,     0),
            ("Clasificacion *",      self.f_clasif, 1),
            ("Nombre Institucion *", self.f_nombre, 2),
            ("RUT",                  self.f_rut,    3),
            ("Cod. Previred *",      self.f_cod,    4),
        ]:
            lbl = QLabel(lbl_text)
            lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            grid.addWidget(lbl, row, 0); grid.addWidget(widget, row, 1)
        main.addWidget(grp)
        nota = QLabel("  * Campos obligatorios")
        nota.setStyleSheet("color:#6b7280;font-size:11px;"); main.addWidget(nota)
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.button(QDialogButtonBox.StandardButton.Ok).setText("Guardar")
        btn_box.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        btn_box.accepted.connect(self._validate_and_accept)
        btn_box.rejected.connect(self.reject)
        main.addWidget(btn_box)

    def _fill_data(self):
        r = self.record
        self.f_id.setText(str(r[0] or ""))
        self.f_clasif.setText(str(r[1] or ""))
        self.f_nombre.setText(str(r[2] or ""))
        self.f_rut.setText(str(r[3] or ""))
        self.f_cod.setText(str(r[4] or ""))

    def _validate_and_accept(self):
        errors = []
        if not self.f_id.text().strip():
            errors.append("- ID APV: campo obligatorio")
        if not self.f_clasif.text().strip():
            errors.append("- Clasificacion: campo obligatorio")
        if not self.f_nombre.text().strip():
            errors.append("- Nombre Institucion: campo obligatorio")
        if not self.f_cod.text().strip():
            errors.append("- Cod. Previred: campo obligatorio")
        if errors:
            QMessageBox.warning(self, "Errores de validacion", "Corrija los siguientes errores:\n\n" + "\n".join(errors))
            return
        self.accept()

    def get_values(self):
        return [
            self.f_id.text().strip(),
            self.f_clasif.text().strip(),
            self.f_nombre.text().strip(),
            self.f_rut.text().strip(),
            self.f_cod.text().strip(),
        ]

class SaludDialog(QDialog):

    def __init__(self, parent=None, record=None):

        super().__init__(parent)

        self.is_edit = record is not None

        self.record  = record

        self.setWindowTitle('Editar Institucion Salud' if self.is_edit else 'Nueva Institucion Salud')

        self.setMinimumWidth(580); self.setModal(True)

        self._build_ui()

        if self.is_edit:

            self._fill_data()



    def _build_ui(self):

        main = QVBoxLayout(self)

        main.setSpacing(12); main.setContentsMargins(20, 16, 20, 12)

        title = QLabel('Editar Institucion Salud' if self.is_edit else 'Nueva Institucion Salud')

        title.setFont(QFont('Segoe UI', 14, QFont.Weight.Bold))

        title.setStyleSheet('color:#1e3a5f;'); main.addWidget(title)

        grp = QGroupBox('Datos de la Institucion de Salud')

        grp.setStyleSheet(GRP_BLUE)

        grid = QGridLayout(grp)

        grid.setSpacing(10)

        grid.setColumnMinimumWidth(0, 160)

        grid.setColumnMinimumWidth(1, 320)

        self.f_id     = make_line('ej: banmedica', 30)

        self.f_clasif = make_line('ej: is', 10)

        self.f_nombre = make_line('ej: ISAPRE Banmedica', 80)

        self.f_rut    = make_line('ej: 96572800-7', 20)

        if self.is_edit:

            self.f_id.setReadOnly(True)

            self.f_id.setStyleSheet('background:#eef2f7;color:#374151;border:1px solid #cbd5e1;border-radius:5px;padding:4px 8px;')

        cod_container = QWidget()

        cod_layout    = QHBoxLayout(cod_container)

        cod_layout.setContentsMargins(0, 0, 0, 0); cod_layout.setSpacing(8)

        self.f_equiv = make_line('ej: 01', 5)

        self.f_equiv.setMaximumWidth(90)

        self.lbl_prev_status = QLabel('')

        self.lbl_prev_status.setMinimumWidth(220)

        cod_layout.addWidget(self.f_equiv)

        cod_layout.addWidget(self.lbl_prev_status)

        cod_layout.addStretch()

        self.f_equiv.textChanged.connect(self._check_previred)

        for lbl_text, widget, row in [

            ('ID Institucion *',      self.f_id,       0),

            ('Clasificacion *',       self.f_clasif,   1),

            ('Nombre Institucion *',  self.f_nombre,   2),

            ('RUT Institucion',       self.f_rut,      3),

            ('Equiv. Previred *',     cod_container,   4),

        ]:

            lbl = QLabel(lbl_text)

            lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

            grid.addWidget(lbl, row, 0); grid.addWidget(widget, row, 1)

        prev_list = '  |  '.join(f'{cod} = {nombre}' for cod, nombre in sorted(PREVIRED_SALUD.items()))

        lbl_ref = QLabel(f'Codigos Previred Salud validos:\n{prev_list}')

        lbl_ref.setWordWrap(True)

        lbl_ref.setStyleSheet('color:#1d4ed8; font-size:11px; background:#eff6ff; border:1px solid #bfdbfe; border-radius:5px; padding:6px 8px;')

        grid.addWidget(lbl_ref, 5, 0, 1, 2)

        main.addWidget(grp)

        nota = QLabel('  * Campos obligatorios')

        nota.setStyleSheet('color:#6b7280;font-size:11px;'); main.addWidget(nota)

        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)

        btn_box.button(QDialogButtonBox.StandardButton.Ok).setText('Guardar')

        btn_box.button(QDialogButtonBox.StandardButton.Cancel).setText('Cancelar')

        btn_box.accepted.connect(self._validate_and_accept)

        btn_box.rejected.connect(self.reject)

        main.addWidget(btn_box)



    def _check_previred(self, text):

        cod = text.strip().zfill(2) if text.strip() else ''

        if not cod:

            self.lbl_prev_status.setText(''); return

        nombre = PREVIRED_SALUD.get(cod)

        if nombre:

            self.lbl_prev_status.setText(f'OK  {cod}  ->  {nombre}')

            self.lbl_prev_status.setStyleSheet('color:#16a34a; font-weight:bold;')

        else:

            self.lbl_prev_status.setText(f'NO  codigo {cod} no existe')

            self.lbl_prev_status.setStyleSheet('color:#dc2626; font-weight:bold;')



    def _fill_data(self):

        r = self.record

        self.f_id.setText(str(r[0] or ''))

        self.f_clasif.setText(str(r[1] or ''))

        self.f_nombre.setText(str(r[2] or ''))

        self.f_rut.setText(str(r[3] or ''))

        self.f_equiv.setText(str(r[4] or ''))

        self._check_previred(self.f_equiv.text())



    def _validate_and_accept(self):

        errors = []

        if not self.f_id.text().strip():

            errors.append('- ID Institucion: campo obligatorio')

        if not self.f_clasif.text().strip():

            errors.append('- Clasificacion: campo obligatorio')

        if not self.f_nombre.text().strip():

            errors.append('- Nombre Institucion: campo obligatorio')

        cod = self.f_equiv.text().strip().zfill(2) if self.f_equiv.text().strip() else ''

        if not cod:

            errors.append('- Equiv. Previred: campo obligatorio')

        elif cod not in PREVIRED_SALUD:

            errors.append(f'- Equiv. Previred {cod}: no existe. Validos: {", ".join(sorted(PREVIRED_SALUD.keys()))}')

        if errors:

            QMessageBox.warning(self, 'Errores de validacion', 'Corrija los siguientes errores:\n\n' + '\n'.join(errors))

            return

        self.accept()



    def get_values(self):

        cod = self.f_equiv.text().strip().zfill(2)

        return [

            self.f_id.text().strip(),

            self.f_clasif.text().strip(),

            self.f_nombre.text().strip(),

            self.f_rut.text().strip(),

            cod,

        ]


# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO LIQUIDACIONES — Parser PDF/Excel + Pestaña UI  (versión 2)
# Mejoras:
#   • Extracción PDF mejorada: más patrones, detección de sección haber/descuento
#   • Detección columnas Excel mejorada: más aliases, normalización agresiva
#   • Historial con filtros (período, RUT, nombre) y paginación
#   • Botón eliminar registros del historial
#   • Validación RUT contra tabla_empleados
#   • Panel de detalle de haberes/descuentos al hacer clic en fila
#   • Cruce con tabla_conceptos para enriquecer conceptos
# ══════════════════════════════════════════════════════════════════════════════

SQL_LIQ_TABLA = """
CREATE TABLE IF NOT EXISTS liquidaciones (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha_carga      TEXT NOT NULL,
    origen           TEXT NOT NULL,
    archivo          TEXT NOT NULL,
    mes_periodo      TEXT,
    rut_trabajador   TEXT,
    nombre           TEXT,
    sueldo_base      REAL,
    total_haberes    REAL,
    total_descuentos REAL,
    liquido_pagar    REAL,
    afp_monto        REAL,
    salud_monto      REAL,
    datos_raw        TEXT
);"""

SQL_LIQ_DETALLE = """
CREATE TABLE IF NOT EXISTS liquidaciones_detalle (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    liquidacion_id INTEGER NOT NULL,
    tipo           TEXT,
    concepto       TEXT,
    monto          REAL,
    FOREIGN KEY (liquidacion_id) REFERENCES liquidaciones(id)
);"""

# ── Helpers BD liquidaciones ───────────────────────────────────────────────────

def _empleados_ruts():
    """Retorna un dict {rut_normalizado: nombre_completo} desde tabla_empleados."""
    try:
        with sqlite3.connect(DB_PATH) as con:
            rows = con.execute("SELECT * FROM tabla_empleados LIMIT 1").fetchall()
            if not rows:
                return {}
            # Detectar columnas: buscamos columna con 'rut' y columna con 'nombre'
            desc = [d[1].lower() for d in con.execute("PRAGMA table_info(tabla_empleados)").fetchall()]
            # desc es lista de nombres de columna
            cols = [d[1] for d in con.execute("PRAGMA table_info(tabla_empleados)").fetchall()]
            rut_col = next((c for c in cols if 'rut' in c.lower()), None)
            nom_cols = [c for c in cols if any(k in c.lower() for k in ('nombre','apellido'))]
            if not rut_col:
                return {}
            sel = f"SELECT {rut_col}" + (f", {', '.join(nom_cols)}" if nom_cols else "") + " FROM tabla_empleados"
            resultado = {}
            for row in con.execute(sel).fetchall():
                r = str(row[0]).strip().upper().replace('.','').replace('-','')
                nombre = ' '.join(str(v) for v in row[1:] if v) if len(row) > 1 else ''
                resultado[r] = nombre
            return resultado
    except Exception:
        return {}


def _conceptos_dict():
    """Retorna dict {codigo_o_nombre_lower: descripcion} desde tabla_conceptos."""
    try:
        with sqlite3.connect(DB_PATH) as con:
            cols = [d[1] for d in con.execute("PRAGMA table_info(tabla_conceptos)").fetchall()]
            if not cols:
                return {}
            # Buscar columna código/id y columna descripción/nombre
            cod_col  = next((c for c in cols if any(k in c.lower() for k in ('cod','id','clave'))), cols[0])
            desc_col = next((c for c in cols if any(k in c.lower() for k in ('desc','nombre','glosa','concepto'))), cols[-1])
            resultado = {}
            for row in con.execute(f"SELECT {cod_col}, {desc_col} FROM tabla_conceptos").fetchall():
                if row[0]:
                    resultado[str(row[0]).strip().lower()] = str(row[1] or '').strip()
            return resultado
    except Exception:
        return {}


def _normalizar_rut(rut):
    """Normaliza RUT a formato SIN puntos ni guión, mayúsculas."""
    return str(rut).strip().upper().replace('.','').replace('-','').replace(' ','')


def _lr_norm_rut(rut):
    """Normaliza RUT: sin puntos, con guión, mayúsculas."""
    r = str(rut).strip().upper().replace('.', '').replace(' ', '')
    if '-' not in r and len(r) > 1:
        r = r[:-1] + '-' + r[-1]
    return r


def liq_historial_fetch(filtro_periodo='', filtro_rut='', filtro_nombre='', limit=500):
    with sqlite3.connect(DB_PATH) as con:
        conds, params = [], []
        if filtro_periodo:
            conds.append("mes_periodo LIKE ?"); params.append(f"%{filtro_periodo}%")
        if filtro_rut:
            conds.append("rut_trabajador LIKE ?"); params.append(f"%{filtro_rut}%")
        if filtro_nombre:
            conds.append("nombre LIKE ?"); params.append(f"%{filtro_nombre}%")
        where = ("WHERE " + " AND ".join(conds)) if conds else ""
        sql = f"""SELECT id,fecha_carga,origen,archivo,mes_periodo,rut_trabajador,nombre,
                         sueldo_base,total_haberes,total_descuentos,liquido_pagar,afp_monto,salud_monto
                  FROM liquidaciones {where} ORDER BY id DESC LIMIT {limit}"""
        return con.execute(sql, params).fetchall()


def liq_detalle_fetch(liquidacion_id):
    with sqlite3.connect(DB_PATH) as con:
        return con.execute(
            "SELECT tipo,concepto,monto FROM liquidaciones_detalle WHERE liquidacion_id=? ORDER BY tipo,rowid",
            (liquidacion_id,)
        ).fetchall()


def liq_delete(ids):
    """Elimina registros de liquidaciones (y sus detalles) por lista de IDs."""
    with sqlite3.connect(DB_PATH) as con:
        for liq_id in ids:
            con.execute("DELETE FROM liquidaciones_detalle WHERE liquidacion_id=?", (liq_id,))
            con.execute("DELETE FROM liquidaciones WHERE id=?", (liq_id,))
        con.commit()


# ── Parser PDF ─────────────────────────────────────────────────────────────────

class _ParserPDF:
    # Patrones mejorados: cubren más variaciones de formato
    PATRONES = {
        "rut": (
            r"R\.?U\.?T\.?\s*(?:N[°º]?\s*)?[:=]?\s*([\d\.]+-[\dkK])",
            r"(?:^|\s)([\d]{7,8}-[\dkK])(?:\s|$)",          # RUT sin prefijo
        ),
        "nombre": (
            r"(?:Nombre\s+(?:Completo|Trabajador)?|Trabajador|Sr\.?a?|Empleado)\s*[:=]?\s*([A-ZÁÉÍÓÚÑ][a-zA-ZÁÉÍÓÚÑáéíóúñ\s,\.]{3,60}?)(?:\n|RUT|Per)",
            r"(?:Nombre|NOMBRE)\s*[:\-]?\s*([A-ZÁÉÍÓÚÑ][a-zA-ZÁÉÍÓÚÑáéíóúñ\s]{3,50})",
        ),
        "periodo": (
            r"(?:Per[ií]odo|PERIODO|Mes\s+Proceso|Liquidaci[oó]n\s+(?:de\s+)?Sueldo)\s*[:\-]?\s*(\w+\s+\d{4}|\d{2}[/-]\d{4}|\d{4}[-/]\d{2})",
            r"(?:Mes|MES)\s*[:\-]?\s*(\w+\s+\d{4}|\d{2}[/-]\d{4})",
        ),
        "sueldo_base": (
            r"(?:Sueldo\s+Base|SUELDO\s+BASE|Remuneraci[oó]n\s+Base|Haber\s+Base)\s*\$?\s*([\d\.,]+)",
        ),
        "total_haberes": (
            r"(?:Total\s+Haberes|TOTAL\s+HABERES|Total\s+Imponible|Total\s+Remuneraciones)\s*\$?\s*([\d\.,]+)",
        ),
        "total_descuentos": (
            r"(?:Total\s+Descuentos|TOTAL\s+DESCUENTOS|Total\s+Deducci[oó]n(?:es)?)\s*\$?\s*([\d\.,]+)",
        ),
        "liquido": (
            r"(?:L[ií]quido\s+a\s+Pagar|L[ií]quido\s+Total|LIQUIDO\s+A\s+PAGAR|Total\s+L[ií]quido)\s*\$?\s*([\d\.,]+)",
            r"(?:MONTO\s+A\s+PAGAR|Monto\s+Pagar)\s*\$?\s*([\d\.,]+)",
        ),
        "afp": (
            r"(?:Cotizaci[oó]n\s+AFP|AFP\s+(?:\w+\s+)?Cotizaci[oó]n|Previsi[oó]n\s+(?:\w+\s+)?AFP)\s*\$?\s*([\d\.,]+)",
            r"(?:^|\s)AFP\s+\w+\s+\$?\s*([\d\.,]+)",
            r"AFP\s*[:\-]?\s*\$?\s*([\d\.,]+)",
        ),
        "salud": (
            r"(?:Cotizaci[oó]n\s+Salud|ISAPRE\s+\w+|FONASA)\s*\$?\s*([\d\.,]+)",
            r"(?:Salud|SALUD)\s*[:\-]?\s*\$?\s*([\d\.,]+)",
        ),
    }

    # Palabras clave de sección
    _KW_HABER    = {"haber","remuner","bono","asignac","gratif","hora extra","comis","benefic","aguinaldo","sueldo"}
    _KW_DESCUENTO= {"descuento","deducci","retenci","cotizac","previsi","afp","isapre","fonasa","impuesto","prestamo"}

    @classmethod
    def extraer_texto(cls, ruta):
        # Usar PyMuPDF directamente: produce texto limpio con sangrias correctas
        try:
            doc = fitz.open(ruta)
            texto = "\n".join(p.get_text() for p in doc)
            if texto.strip():
                return texto, "pymupdf"
        except Exception:
            pass
        # Fallback pdfplumber
        if _PDFPLUMBER_OK:
            try:
                with pdfplumber.open(ruta) as pdf:
                    partes = [p.extract_text() or "" for p in pdf.pages]
                    texto = "\n".join(partes)
                if texto.strip():
                    return texto, "pdfplumber"
            except Exception:
                pass
        # Fallback pypdf
        if _PYPDF_OK:
            try:
                reader = PdfReader(ruta)
                texto = "\n".join(p.extract_text() or "" for p in reader.pages)
                if texto.strip():
                    return texto, "pypdf"
            except Exception:
                pass

        # Fallback OCR: rasterizar páginas con PyMuPDF y leer con Tesseract
        if _TESSERACT_OK:
            try:
                if os.path.exists(_TESSERACT_CMD):
                    pytesseract.pytesseract.tesseract_cmd = _TESSERACT_CMD
                doc = fitz.open(ruta)
                partes = []
                for pag in doc:
                    # 300 DPI → calidad óptima para OCR
                    mat  = fitz.Matrix(300 / 72, 300 / 72)
                    pix  = pag.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
                    img  = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    text = pytesseract.image_to_string(img, lang="spa", config="--psm 6")
                    partes.append(text)
                texto = "\n".join(partes)
                if texto.strip():
                    return texto, "ocr-tesseract"
            except Exception as e:
                pass

        return "", "sin_texto"

    @classmethod
    def _limpiar_monto(cls, s):
        """Convierte '1.234.567' o '1234,56' a float."""
        try:
            s = s.strip().replace('$','').replace(' ','')
            # Formato chileno: puntos como miles, coma como decimal
            if ',' in s:
                s = s.replace('.','').replace(',','.')
            else:
                s = s.replace('.','')
            return float(s)
        except Exception:
            return None

    @classmethod
    def _buscar_patron(cls, texto, patrones):
        """Prueba lista de patrones y retorna primer match."""
        for p in patrones:
            m = re.search(p, texto, re.IGNORECASE | re.MULTILINE)
            if m:
                return m.group(1).strip()
        return None

    @classmethod
    def _detectar_seccion(cls, linea_lower):
        if any(k in linea_lower for k in cls._KW_HABER):
            return "HABER"
        if any(k in linea_lower for k in cls._KW_DESCUENTO):
            return "DESCUENTO"
        return None

    # Mapeo nombre OCR → id_concepto en tabla_conceptos
    _MAPA_CONCEPTOS = {
        "sueldo base":                    "sueldoBase",
        "gratificacion legal":            "gratificacion",
        "gratificación legal":            "gratificacion",
        "bono de produccion":             "boProd",
        "bono produccion":                "boProd",
        "comision por ventas":            "comision",
        "comisión por ventas":            "comision",
        "semana corrida":                 "semanaCorr",
        "horas extras 50%":               "horasEx50",
        "horas extras 30%":               "horasEx30",
        "colacion":                       "colacion",
        "colación":                       "colacion",
        "movilizacion":                   "movilizacion",
        "movilización":                   "movilizacion",
        "cargas familiares simples":       "cargasSimp",
        "bono sala cuna":                 "salaCuna",
        "ias vacaciones legales":         "iasVacaciones",
        "totalesempl":                    "totalesEmpl",
        "sueldo liquido":                 "totalesEmpl",
        "sueldo líquido":                 "totalesEmpl",
        "liquido a pagar":                "totalesEmpl",
        "líquido a pagar":                "totalesEmpl",
        "total a pagar":                  "totalesEmpl",
        "total a pago":                   "totalesEmpl",
        "total liquido":                  "totalesEmpl",
        "total líquido":                  "totalesEmpl",
        "total neto":                     "totalesEmpl",
        "liquido":                        "totalesEmpl",
        "líquido":                        "totalesEmpl",
        "neto a pagar":                   "totalesEmpl",
        "monto a pagar":                  "totalesEmpl",
        # Descuentos legales — salud (base + adicional, todos mapean a isapre)
        "adicional isapre":               "isapre",
        "adicional salud":                "isapre",
        "cotizacion salud adicional":     "isapre",
        "cotizacion adicional salud":     "isapre",
        "desc salud adicional":           "isapre",
        "descuento salud adicional":      "isapre",
        "salud adicional":                "isapre",
        "salud":                          "isapre",
        "afp":                            "afp",
        "cotizacion salud":               "isapre",
        "cotizacion afp":                 "afp",
        "impuesto unico":                 "impuesto",
        "impuesto único":                 "impuesto",
        "seguro de cesantia (trabajador": "cesEmpleado",
        "seguro de cesantia":             "cesEmpleado",
        "seguro cesantia":                "cesEmpleado",
        "cesantia trabajador":            "cesEmpleado",
        "seg cesantia":                   "cesEmpleado",
        "seg de cesantia":                "cesEmpleado",
        "cotizacion cesantia":            "cesEmpleado",
        "trabajo pesado empleado":        "trabajoPesaEmpl",
        "apvi ahorro voluntario":         "apvi",
        "anticipo":                       "anticipo",
        "creditos personales ccaf":       "cajaCred",
        "descuentos varios":              "descuentosVarios",
        "retencion judicial":             "retencionJudicial",
        "trabajo pesado":                 "trabajoPesa",
        # Aportes empleador
        "aporte a ccaf":                  "cajaComp",
        "mutual":                         "mutual",
        "seguro invalidez y sobrevivencia":"sis",
        "seguro de cesantia ci":          "cesAporteCi",
        "seguro de cesantia solidario":   "cesAporteSol",
        "aporte afp empleador":           "aporteAFPemp",
        "aporte fapp compensacion":       "aporteFAPPCEV",
        "aporte fapp compensación":       "aporteFAPPCEV",
        # Variantes OCR adicionales
        "cotizacion salud neta":           "isapre",
        # "seguro salud neto" es haber (bonif. salud empresa), NO descuento isapre
        "impuesto":                        "impuesto",
        "impuesto segunda categoria":      "impuesto",
        "imp segunda categoria":           "impuesto",
        "imp 2da categoria":               "impuesto",
        "impuesto 2da categoria":          "impuesto",
        "seguro complementario":           "seguroComp",
        "seguro compl":                    "seguroComp",
        "seguro de invalidez":             "sis",
        "sis":                             "sis",
        "seguro salud neto":               "segSaludNeto",
        "costo empresa":                   "seguroComp",
        "costo colaborador":               "seguroComp",
        "gratificacion":                   "gratificacion",
        "gratificación":                   "gratificacion",
        "asignacion de herramienta":       "asDeHerra",
        "asignación de herramienta":       "asDeHerra",
        "bono turno":                      "boTurno",
        "bono turno nocturno":             "boTurno",
        "bono categoria":                  "boCat",
        "bono categoría":                  "boCat",
    }

    @classmethod
    def _tipo_concepto(cls, id_concepto):
        """Retorna el Tipo desde tabla_conceptos para un id dado."""
        try:
            with sqlite3.connect(DB_PATH) as con:
                row = con.execute(
                    "SELECT Tipo FROM tabla_conceptos WHERE Id=?", (id_concepto,)
                ).fetchone()
                return row[0] if row else ""
        except Exception:
            return ""

    @classmethod
    def _tipo_to_seccion(cls, tipo_bd):
        """Convierte el Tipo de tabla_conceptos a sección interna."""
        t = (tipo_bd or "").lower()
        if "haber afecto" in t or "haber solo imp" in t or "haber solo trib" in t:
            return "HABER_AFECTO"
        if "haber exento" in t:
            return "HABER_EXENTO"
        if "descuento legal" in t:
            return "DESCUENTO_LEGAL"
        if "descuento" in t:
            return "DESCUENTO_OTRO"
        if "aporte" in t:
            return "APORTE"
        return None
    _AFECTO_TOTAL_IMP = {"afp","isapre","trabajoPesaEmpl","mutual","sis","aporteAFPemp","aporteFAPPCEV"}
    _AFECTO_TOPE_CES  = {"cesEmpleado","cesAporteCi","cesAporteSol"}
    _NO_IMPONIBLE     = {"colacion","movilizacion","cargasSimp","iasVacaciones","salaCuna"}

    # Meses en español
    _MESES = {"enero":"01","febrero":"02","marzo":"03","abril":"04","mayo":"05","junio":"06",
              "julio":"07","agosto":"08","septiembre":"09","octubre":"10","noviembre":"11","diciembre":"12"}

    @classmethod
    def _norm_concepto(cls, s):
        import re as _re
        s = s.strip().lower()
        for a,b in [('á','a'),('é','e'),('í','i'),('ó','o'),('ú','u'),('ñ','n')]:
            s = s.replace(a,b)
        # Cortar sufijos OCR: 'cotizacion afp: en AFP Cuprum...' -> 'cotizacion afp'
        s = _re.split(r'[:\(\[]', s)[0].strip()
        # Quitar numeros/porcentajes al final
        s = _re.sub(r'\s+[\d\.,%]+$', '', s).strip()
        return s

    @classmethod
    def _id_concepto(cls, nombre_ocr):
        n = cls._norm_concepto(nombre_ocr)
        # Búsqueda exacta en mapa fijo
        if n in cls._MAPA_CONCEPTOS:
            return cls._MAPA_CONCEPTOS[n], False
        # Búsqueda parcial en mapa fijo
        for k, v in cls._MAPA_CONCEPTOS.items():
            if k in n or n in k:
                return v, False
        # Búsqueda en tabla_conceptos (conceptos agregados por el usuario)
        # IMPORTANTE: solo buscar por nombre exacto para no interferir con ids internos
        try:
            with sqlite3.connect(DB_PATH) as con:
                # Buscar por nombre exacto
                row = con.execute(
                    "SELECT id FROM tabla_conceptos WHERE LOWER(nombre)=LOWER(?)",
                    (nombre_ocr.strip(),)
                ).fetchone()
                if row:
                    return row[0], False
                # Buscar por nombre exacto normalizado (sin tildes, minusculas)
                row = con.execute(
                    "SELECT id FROM tabla_conceptos WHERE LOWER(nombre)=LOWER(?)",
                    (n,)
                ).fetchone()
                if row:
                    return row[0], False
        except Exception:
            pass
        return nombre_ocr, True  # sin mapeo

    @classmethod
    def _parsear_pagina(cls, texto_pagina, ruta):
        """
        Parsea una liquidacion en formato nativo o OCR.
        Formato A (nativo): concepto en linea, monto en linea siguiente con sangria
        Formato B (OCR):    concepto y monto en la misma linea
        """
        txt = texto_pagina

        def buscar(patrones):
            for p in patrones:
                m = re.search(p, txt, re.IGNORECASE)
                if m:
                    return m.group(1).strip()
            return ""

        def monto(s):
            try:
                s = str(s).replace("$","").replace(" ","").replace(".","").replace(",",".").strip()
                return float(s) if s else 0.0
            except Exception:
                return 0.0

        # ── Encabezado ───────────────────────────────────────────────────
        ruts = re.findall(r"([\d\.]{6,}-[\dkK])", txt)
        rut_emp  = _lr_norm_rut(ruts[0]) if len(ruts) > 0 else ""
        rut_trab = _lr_norm_rut(ruts[1]) if len(ruts) > 1 else ""

        nombre_trab = buscar([
            r"Sr\.?\s*([A-ZÁÉÍÓÚÑ][^\n]{5,60?})\s+\d{6,}",
            r"Trabajador[:\s]+([A-ZÁÉÍÓÚÑa-záéíóúñ,\s\.]{5,60?})\s+\d",
        ])

        periodo = ""
        m_per = re.search(
            r"(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)"
            r"[\s\-–]+(\d{4})", txt, re.IGNORECASE)
        if m_per:
            mes_num = cls._MESES.get(m_per.group(1).lower(), "01")
            periodo = f"{m_per.group(2)}-{mes_num}"

        afp_nombre    = buscar([r"en\s+AFP\s+([A-Za-záéíóúñ]+)", r"AFP[:\s]+([^\n,]+?)\s+(?:Porc|\d)"])
        isapre_nombre = buscar([r"en\s+(Consalud|Banmedica|Colmena|Cruz\s+Blanca|Masvida|Esencial|VidaTres)",
                                 r"ISAPRE[:\s]+([^\n]+?)\s+Plan"])
        if not isapre_nombre and re.search(r"fonasa", txt, re.IGNORECASE):
            isapre_nombre = "Fonasa"

        porc_afp_txt = buscar([r"([\d,\.]+)%\s+sobre"])
        try:
            porc_afp = float(str(porc_afp_txt).replace(",","."))
        except Exception:
            porc_afp = 0.0

        # ── Detectar formato ─────────────────────────────────────────────
        lineas = txt.splitlines()
        es_formato_a = sum(
            1 for l in lineas if re.match(r"^\s{4,}[\d\.]{4,}", l)
        ) >= 3

        # ── Palabras a ignorar ───────────────────────────────────────────
        _IGNORAR = {
            "concepto","haberes","descuentos","totales","liquidaci",
            "certifico","recib","forma de pago","abono","banco",
            "tributabl","base imponible","descuentos legales","alcance",
            "plan salud","c. costo","cargo","ingreso","dias trabajados",
            "nro. contrato","tope","utm","uuf","porc",
        }
        _STOP = {
            "totales:","total:","liquido a pagar","líquido a pagar",
            "sueldo liquido","sueldo líquido","certifico haber",
        }

        lineas_detalle = []
        seccion_actual = "HABER_AFECTO"

        if es_formato_a:
            # Formato A: concepto en linea N, monto indentado en linea N+1
            i = 0
            while i < len(lineas):
                ll = lineas[i].strip()
                ll_lower = ll.lower()

                if any(s in ll_lower for s in _STOP):
                    m_liq = re.search(r"([\d\.]{4,}(?:,\d{1,2})?)", ll)
                    if m_liq:
                        lineas_detalle.append(("DESCUENTO_OTRO","Liquido a pagar",monto(m_liq.group(1))))
                    i += 1
                    continue

                if re.search(r"haberes?.*(afect|imponible)", ll_lower):
                    seccion_actual = "HABER_AFECTO"
                elif re.search(r"haberes?.*(exent|no impon)", ll_lower):
                    seccion_actual = "HABER_EXENTO"
                elif re.search(r"\bhaberes?\b", ll_lower) and "total" not in ll_lower:
                    seccion_actual = "HABER_AFECTO"
                elif re.search(r"descuentos?.*(legales?)", ll_lower):
                    seccion_actual = "DESCUENTO_LEGAL"
                elif re.search(r"\bdescuentos?\b", ll_lower) and "total" not in ll_lower:
                    seccion_actual = "DESCUENTO_OTRO"

                if any(ig in ll_lower for ig in _IGNORAR) and not re.search(r"[\d\.]{4,}", ll):
                    i += 1
                    continue

                if ll and i + 1 < len(lineas):
                    # Buscar monto: linea i+1 o i+2 (puede haber linea vacia entre medio)
                    sig = None
                    salto = 0
                    for _k in (1, 2):
                        if i + _k < len(lineas):
                            _cand = lineas[i + _k]
                            if re.match(r"^\s{4,}[\d\.]{3,}", _cand):
                                sig = _cand
                                salto = _k
                                break
                    if sig is not None:
                        v = monto(sig.strip())
                        if v > 0 and len(ll) >= 3:
                            nombre_limpio = re.split(r"[:\(]", ll)[0].strip()
                            if not any(ig in nombre_limpio.lower() for ig in _IGNORAR):
                                lineas_detalle.append((seccion_actual, nombre_limpio, v))
                        i += salto + 1
                        continue
                i += 1

        else:
            # Formato B (OCR): concepto y monto en la misma linea
            patron = re.compile(
                r"([\w\s\(\)áéíóúñÁÉÍÓÚÑ,\./-]{3,50}?)\s+([\d\.]{3,}(?:,\d{1,2})?)"
            )
            for linea in lineas:
                ll = linea.strip()
                ll_lower = ll.lower()
                if any(s in ll_lower for s in _STOP):
                    m_liq = re.search(r"([\d\.]{5,}(?:,\d{1,2})?)\s*$", ll)
                    if m_liq:
                        lineas_detalle.append(("DESCUENTO_OTRO","Liquido a pagar",monto(m_liq.group(1))))
                    continue
                if re.search(r"haberes?.*(afect|imponible)", ll_lower):
                    seccion_actual = "HABER_AFECTO"
                elif re.search(r"descuentos?.*(legales?)", ll_lower):
                    seccion_actual = "DESCUENTO_LEGAL"
                elif re.search(r"\bdescuentos?\b", ll_lower) and "total" not in ll_lower:
                    seccion_actual = "DESCUENTO_OTRO"

                matches = list(patron.finditer(ll))
                for i2, m2 in enumerate(matches):
                    nombre = re.split(r"[:\(]", m2.group(1).strip())[0].strip()
                    v = monto(m2.group(2))
                    if v <= 0 or len(nombre) < 3:
                        continue
                    if re.match(r"^[\d\.\s,\(\)\-]+$", nombre):
                        continue
                    if any(ig in nombre.lower() for ig in _IGNORAR):
                        continue
                    sec = seccion_actual if len(matches) == 1 else (
                        "HABER_AFECTO" if i2 == 0 else "DESCUENTO_LEGAL"
                    )
                    lineas_detalle.append((sec, nombre, v))

        # ── Deduplicar por id_concepto ────────────────────────────────────
        _vistos = {}
        lineas_f2 = []
        for _sec, _nom, _v in lineas_detalle:
            _id, _sm = cls._id_concepto(_nom)
            if re.match(r"^[\d\.\s,\(\)\-]+$", _nom.strip()):
                continue
            if _id not in _vistos:
                _vistos[_id] = True
                lineas_f2.append((_sec, _nom, _v))
            elif _sm:
                lineas_f2.append((_sec, _nom, _v))
        lineas_detalle = lineas_f2

        # ── Fusionar salud ────────────────────────────────────────────────
        lineas_detalle_f = []
        isapre_ac = 0.0; isapre_sec = "DESCUENTO_LEGAL"
        for sec, nom, v in lineas_detalle:
            if cls._id_concepto(nom)[0] == "isapre":
                isapre_ac += v; isapre_sec = sec
            else:
                lineas_detalle_f.append((sec, nom, v))
        if isapre_ac > 0:
            lineas_detalle_f.append((isapre_sec, "Salud", isapre_ac))
        lineas_detalle = lineas_detalle_f

        # ── Calcular totales para Afecto ─────────────────────────────────
        empresa_id = _lr_get_empresa_id(rut_emp)
        try:
            with sqlite3.connect(DB_PATH) as _ce:
                _re = _ce.execute("SELECT 1 FROM empresas WHERE rut_emp=?", (rut_emp,)).fetchone()
                empresa_no_bd = (_re is None and bool(rut_emp))
        except Exception:
            empresa_no_bd = False

        n_contrato  = _lr_get_numcontrato(rut_trab)
        cot_mutual  = _lr_get_cot_mutual(rut_emp)
        inst_afp_id = _lr_get_inst_id(afp_nombre, "afp")
        inst_sal_id = _lr_get_inst_id(isapre_nombre, "salud")

        tope_imp_bd = _lr_get_tope_imp(periodo) if periodo else 0.0
        tope_ces_bd = _lr_get_tope_ces(periodo) if periodo else 0.0
        tope_sal_bd = _lr_get_tope_salud(periodo) if periodo else 0.0

        monto_afp   = next((v for s,n,v in lineas_detalle if cls._id_concepto(n)[0]=="afp"), 0)
        monto_salud = next((v for s,n,v in lineas_detalle if cls._id_concepto(n)[0]=="isapre"), 0)

        # Si no se detectaron secciones por encabezado, inferir desde el id del concepto
        _ids_descuento = cls._AFECTO_TOTAL_IMP | cls._AFECTO_TOPE_CES | {
            "impuesto","cesEmpleado","isapre","mutual","trabajoPesa","trabajoPesaEmpl",
            "apvi","anticipo","cajaCred","descuentosVarios","retencionJudicial",
            "seguroComp","totalesEmpl",
        }
        hay_haber_afecto = any(s == "HABER_AFECTO" for s,n,v in lineas_detalle)
        lineas_detalle_corr = []
        for _s, _n, _v in lineas_detalle:
            _id, _ = cls._id_concepto(_n)
            if not hay_haber_afecto:
                # Inferir seccion desde el id
                if _id in cls._AFECTO_TOTAL_IMP or _id in cls._AFECTO_TOPE_CES or _id == "impuesto":
                    _s = "DESCUENTO_LEGAL"
                elif _id in ("totalesEmpl","seguroComp","anticipo","cajaCred",
                             "descuentosVarios","retencionJudicial","apvi"):
                    _s = "DESCUENTO_OTRO"
                elif _id not in _ids_descuento:
                    _s = "HABER_AFECTO" if _id not in cls._NO_IMPONIBLE else "HABER_EXENTO"
            lineas_detalle_corr.append((_s, _n, _v))
        lineas_detalle = lineas_detalle_corr

        # Suma bruta de haberes imponibles (sin topes aun)
        afecto_bruto = sum(v for s,n,v in lineas_detalle
                           if s in ("HABER_AFECTO","haber")
                           and cls._id_concepto(n)[0] not in cls._NO_IMPONIBLE
                           and cls._id_concepto(n)[0] not in cls._AFECTO_TOTAL_IMP
                           and cls._id_concepto(n)[0] not in cls._AFECTO_TOPE_CES
                           and cls._id_concepto(n)[0] != "impuesto")

        # Tope AFP (independiente del tope cesantia)
        afecto_imp_base = min(afecto_bruto, tope_imp_bd) if tope_imp_bd else afecto_bruto

        # Tope cesantia se aplica sobre el bruto, NO sobre el ya limitado por AFP
        afecto_ces  = min(afecto_bruto, tope_ces_bd) if tope_ces_bd else afecto_bruto
        afecto_imp  = afecto_imp_base
        salud_base  = min(monto_salud, tope_sal_bd) if tope_sal_bd else monto_salud
        desc_legales = monto_afp + salud_base + afecto_ces * 0.006

        # ── Generar filas ────────────────────────────────────────────────
        filas = []
        for seccion, nombre_orig, v in lineas_detalle:
            id_conc, sin_mapeo = cls._id_concepto(nombre_orig)

            if id_conc in cls._AFECTO_TOTAL_IMP:
                afecto = int(afecto_imp_base)
            elif id_conc in cls._AFECTO_TOPE_CES:
                afecto = int(afecto_ces)
            elif id_conc == "impuesto":
                afecto = int(afecto_imp)
            else:
                afecto = 0

            if id_conc == "afp":
                id_inst = inst_afp_id
            elif id_conc == "isapre":
                id_inst = inst_sal_id
            elif id_conc in ("cesEmpleado","cesAporteCi","cesAporteSol","sis","aporteAFPemp","aporteFAPPCEV"):
                id_inst = inst_afp_id
            else:
                id_inst = ""

            if id_conc == "afp":
                cot_jub = porc_afp
            elif id_conc == "isapre":
                cot_jub = monto_salud
            elif id_conc == "mutual":
                cot_jub = cot_mutual
            elif id_conc == "cesEmpleado":
                cot_jub = 0.6
            else:
                cot_jub = 0

            rebajas_llss = int(desc_legales) if id_conc == "impuesto" else 0

            filas.append({
                "Fecha de proceso":          periodo,
                "Id empleado":               rut_trab,
                "Número de contrato":        n_contrato,
                "Id del concepto":           id_conc,
                "Monto del concepto":        int(v),
                "Afecto":                    afecto,
                "Id de institución":         id_inst,
                "Cotización de jubilación":  cot_jub,
                "Días de licencias":         0,
                "Días trabajados":           30,
                "Fecha de aplicación":       "x",
                "Empresa":                   empresa_id,
                "Total de rebajas por LLSS": rebajas_llss,
                "Rentas no gravadas":        0,
                "Rebaja por zona extrema":   0,
                "Jornada":                   "C",
                "_col_origen":               nombre_orig,
                "_sin_mapeo":                sin_mapeo,
                "_seccion":                  seccion,
                "_nombre_trab":              nombre_trab,
                "_rut_emp":                  rut_emp,
                "_empresa_no_bd":            empresa_no_bd,
                "_nombre_empresa":           buscar([r"([A-Z][A-Za-z\s\.]+(?:S\.A\.|SPA|LTDA|EIRL|SA))"]),
            })

        # ── Aportes empleador ────────────────────────────────────────────
        afecto_afp_val  = next((f["Afecto"] for f in filas if f["Id del concepto"]=="afp"), 0)
        afecto_ces_val  = next((f["Afecto"] for f in filas if f["Id del concepto"]=="cesEmpleado"), 0)
        if not afecto_ces_val and afecto_afp_val:
            afecto_ces_val = min(afecto_afp_val, int(_lr_get_tope_ces(periodo) or afecto_afp_val))
        id_inst_afp_val = next((f["Id de institución"] for f in filas if f["Id del concepto"]=="afp"), "")

        aportes = _lr_calcular_aportes(
            rut_trab, periodo, afecto_afp_val, afecto_ces_val,
            id_inst_afp_val, id_inst_afp_val
        )
        for ap in aportes:
            ap.update({
                "Fecha de proceso":          periodo,
                "Id empleado":               rut_trab,
                "Número de contrato":        n_contrato,
                "Días de licencias":         0,
                "Días trabajados":           30,
                "Fecha de aplicación":       "x",
                "Empresa":                   empresa_id,
                "Total de rebajas por LLSS": 0,
                "Rentas no gravadas":        0,
                "Rebaja por zona extrema":   0,
                "Jornada":                   "C",
                "_col_origen":               ap["Id del concepto"],
                "_sin_mapeo":                False,
                "_seccion":                  "APORTE",
                "_nombre_trab":              nombre_trab,
                "_rut_emp":                  rut_emp,
                "_empresa_no_bd":            empresa_no_bd,
                "_nombre_empresa":           buscar([r"([A-Z][A-Za-z\s\.]+(?:S\.A\.|SPA|LTDA|EIRL|SA))"]),
            })
        filas.extend(aportes)
        return filas


    @classmethod
    def parsear(cls, ruta):
        partes  = []
        metodo  = "sin_texto"

        # Estrategia 1: texto nativo (PDFs digitales) — más limpio y preciso
        try:
            texto_nativo, _ = cls.extraer_texto(ruta)
            if texto_nativo and texto_nativo.strip():
                separador = re.compile(
                    r"(?=LIQUIDACI[OÓ]N\s+DE\s+REMUNERACIONES)", re.IGNORECASE)
                partes = [p.strip() for p in separador.split(texto_nativo) if p.strip()]
                if not partes:
                    partes = [texto_nativo.strip()]
                metodo = "nativo"
        except Exception:
            pass

        # Estrategia 2: OCR con Tesseract (PDFs escaneados sin texto nativo)
        if not partes and _TESSERACT_OK:
            try:
                if os.path.exists(_TESSERACT_CMD):
                    pytesseract.pytesseract.tesseract_cmd = _TESSERACT_CMD
                doc = fitz.open(ruta)
                for pag in doc:
                    mat = fitz.Matrix(300 / 72, 300 / 72)
                    pix = pag.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    txt_pag = pytesseract.image_to_string(img, lang="spa", config="--psm 6")
                    if txt_pag.strip():
                        partes.append(txt_pag.strip())
                if partes:
                    metodo = "ocr-tesseract"
            except Exception:
                pass

        # Fallback final: texto nativo si OCR también falló
        if not partes:
            texto, metodo = cls.extraer_texto(ruta)
            if texto.strip():
                separador = re.compile(
                    r"(?=LIQUIDACI[OÓ]N\s+DE\s+REMUNERACIONES)", re.IGNORECASE)
                partes = [p.strip() for p in separador.split(texto) if p.strip()]
                if not partes:
                    partes = [texto]

        res_base = {
            "origen": "PDF", "archivo": os.path.basename(ruta),
            "metodo": metodo, "texto_crudo": "",
            "advertencias": [], "detalle": [],
        }
        if metodo == "ocr-tesseract":
            res_base["advertencias"].append("ℹ OCR aplicado con Tesseract (PDF escaneado).")
        if not partes:
            res_base["advertencias"].append(
                "⚠ No se extrajo texto. Instala pytesseract y Tesseract-OCR.")
            return [res_base]

        resultados = []
        for parte in partes:
            try:
                filas = cls._parsear_pagina(parte, ruta)
                if filas:
                    # Para compatibilidad con el worker, empaquetar como un resultado
                    # con las filas en "_filas_libro" para que el exportador las use
                    r = dict(res_base)
                    r["texto_crudo"] = parte
                    r["_filas_libro"] = filas
                    r["rut_trabajador"] = filas[0]["Id empleado"] if filas else ""
                    r["nombre"] = filas[0].get("_nombre_trab","") if filas else ""
                    r["mes_periodo"] = filas[0]["Fecha de proceso"] if filas else ""
                    # Extraer totales para la vista de la tabla
                    haberes = sum(f["Monto del concepto"] for f in filas
                                  if any(f["Id del concepto"]==c for c in
                                         ["sueldoBase","gratificacion","semanaCorr","boProd","comision"]))
                    r["total_haberes"]    = haberes
                    r["liquido_pagar"]    = next((f["Monto del concepto"] for f in filas
                                                  if "liquido" in f["_col_origen"].lower()), 0)
                    r["advertencias"]     = list(res_base["advertencias"])
                    sin_mapeo = [f["_col_origen"] for f in filas if f.get("_sin_mapeo")]
                    if sin_mapeo:
                        r["advertencias"].append(
                            f"⚠ Sin mapeo: {', '.join(set(sin_mapeo))}")
                    resultados.append(r)
            except Exception as e:
                r = dict(res_base)
                r["advertencias"] = [f"❌ Error parseando liquidación: {e}"]
                resultados.append(r)

        return resultados if resultados else [res_base]


# ── Parser Excel ───────────────────────────────────────────────────────────────

class _ParserExcel:
    # Aliases ampliados y normalizados (sin tildes, minúsculas)
    ALIAS = {
        "rut_trabajador": [
            "rut","r.u.t","rut trabajador","rut trabajador","run",
            "n° rut","nro rut","rut empleado","rut_trabajador",
        ],
        "nombre": [
            "nombre","nombre trabajador","trabajador","empleado",
            "apellidos y nombre","nombre completo","nombre y apellido",
            "apellido paterno","apellido","nombres",
        ],
        "mes_periodo": [
            "periodo","período","mes","mes proceso","mes periodo",
            "mes/año","mes año","periodo remuneracion","periodo liquidacion",
        ],
        "sueldo_base": [
            "sueldo base","sueldo","remuneracion base","remuneración base",
            "sueldo bruto","haber base","remuneracion mensual",
            "salario base","salario","sueldo mensual",
        ],
        "total_haberes": [
            "total haberes","total imponible","total remuneraciones",
            "total haberes imponibles","suma haberes","total bruto",
            "total haber",
        ],
        "total_descuentos": [
            "total descuentos","total deducciones","total deduccion",
            "suma descuentos","descuentos totales","total descuento",
        ],
        "liquido_pagar": [
            "liquido a pagar","líquido a pagar","liquido","liquido total",
            "total liquido","monto a pagar","neto a pagar",
            "total neto","neto","líquido",
            "total a pagar","total a pago","sueldo liquido","sueldo líquido",
        ],
        "afp_monto": [
            "afp","prevision","previsión","cotizacion afp","cotización afp",
            "descuento afp","monto afp","afp obligatoria",
        ],
        "salud_monto": [
            "salud","isapre","fonasa","cotizacion salud","cotización salud",
            "descuento salud","monto salud","prevision salud",
        ],
    }

    @classmethod
    def _norm(cls, s):
        """Normaliza texto: minúsculas, sin tildes, sin espacios extra."""
        if not s:
            return ""
        s = str(s).strip().lower()
        for a, b in [('á','a'),('é','e'),('í','i'),('ó','o'),('ú','u'),('ü','u'),('ñ','n')]:
            s = s.replace(a, b)
        return re.sub(r'\s+', ' ', s)

    @classmethod
    def listar_hojas(cls, ruta):
        try:
            wb = load_workbook(ruta, read_only=True, data_only=True)
            return wb.sheetnames
        except Exception:
            return []

    @classmethod
    def parsear(cls, ruta, hoja=None):
        wb = load_workbook(ruta, data_only=True)
        ws = wb[hoja] if hoja and hoja in wb.sheetnames else wb.active
        filas = list(ws.iter_rows(values_only=True))
        if not filas:
            return []

        # Tabla normalizada de aliases para búsqueda rápida
        todos_alias_norm = {cls._norm(a): (campo, a)
                            for campo, lista in cls.ALIAS.items() for a in lista}

        # Detectar fila de encabezados (primera que tenga ≥2 coincidencias conocidas)
        idx_hdr = 0
        mejor_score = 0
        for i, fila in enumerate(filas[:30]):
            score = sum(1 for c in fila if cls._norm(c) in todos_alias_norm)
            if score > mejor_score:
                mejor_score = score
                idx_hdr = i
            if score >= 3:
                break

        headers_raw  = filas[idx_hdr]
        headers_norm = [cls._norm(c) for c in headers_raw]

        # Mapear campo → índice de columna
        mapa = {}
        for i, hn in enumerate(headers_norm):
            if hn in todos_alias_norm:
                campo = todos_alias_norm[hn][0]
                if campo not in mapa:   # primer match gana
                    mapa[campo] = i

        resultados = []
        for fila in filas[idx_hdr + 1:]:
            if all(c is None for c in fila):
                continue
            reg = {
                "origen": "EXCEL", "archivo": os.path.basename(ruta),
                "advertencias": [], "detalle": [],
            }
            for campo, idx in mapa.items():
                if idx < len(fila) and fila[idx] is not None:
                    reg[campo] = fila[idx]

            # Construir raw_extra con columnas no mapeadas
            mapa_vals = set(mapa.values())
            raw = {}
            for i, v in enumerate(fila):
                if v is not None and i not in mapa_vals:
                    key = str(headers_raw[i]) if i < len(headers_raw) else f"col_{i}"
                    raw[key] = v
            reg["raw_extra"] = raw

            # Intentar construir detalle desde raw_extra
            # Si hay columnas que parezcan haberes/descuentos individuales con montos
            for col_name, val in raw.items():
                cn = cls._norm(col_name)
                if isinstance(val, (int, float)) and val != 0:
                    tipo = "HABER" if any(k in cn for k in ("haber","bono","asig","gratif","comis")) \
                           else "DESCUENTO" if any(k in cn for k in ("desc","afp","salud","isapre","fonasa","cotiz")) \
                           else None
                    if tipo:
                        reg["detalle"].append({"tipo": tipo, "concepto": str(col_name), "monto": float(val)})

            if reg.get("rut_trabajador") or reg.get("nombre"):
                # Advertir campos faltantes
                for fld in ("liquido_pagar","total_haberes"):
                    if not reg.get(fld):
                        reg["advertencias"].append(f"⚠ Sin campo: {fld}")
                resultados.append(reg)
        return resultados


# ── Worker thread ──────────────────────────────────────────────────────────────

class _WorkerLiq(QThread):
    progreso  = pyqtSignal(int, str)
    terminado = pyqtSignal(list)

    def __init__(self, archivos, hoja=None):
        super().__init__()
        self.archivos = archivos
        self.hoja = hoja

    def run(self):
        # Cargar tablas de referencia una sola vez
        empleados  = _empleados_ruts()     # {rut_norm: nombre}
        conceptos  = _conceptos_dict()     # {cod_lower: descripcion}

        resultados = []
        total = len(self.archivos)
        for i, ruta in enumerate(self.archivos):
            self.progreso.emit(int(i / total * 100), os.path.basename(ruta))
            try:
                ext = os.path.splitext(ruta)[1].lower()
                if ext == ".pdf":
                    parsed = _ParserPDF.parsear(ruta)
                    # parsear ahora retorna lista de resultados (uno por liquidación/página)
                    if isinstance(parsed, list):
                        for r in parsed:
                            self._enriquecer(r, empleados, conceptos)
                        resultados.extend(parsed)
                    else:
                        self._enriquecer(parsed, empleados, conceptos)
                        resultados.append(parsed)
                elif ext in (".xlsx", ".xls", ".xlsm"):
                    lista = _ParserExcel.parsear(ruta, self.hoja)
                    for res in lista:
                        self._enriquecer(res, empleados, conceptos)
                    resultados.extend(lista)
            except Exception as e:
                resultados.append({"origen": "ERROR", "archivo": os.path.basename(ruta), "error": str(e)})
        self.progreso.emit(100, "Listo")
        self.terminado.emit(resultados)

    @staticmethod
    def _enriquecer(reg, empleados, conceptos):
        """Valida RUT contra tabla_empleados y enriquece conceptos de detalle."""
        rut_raw = str(reg.get("rut_trabajador","")).strip()
        if rut_raw:
            rut_norm = _normalizar_rut(rut_raw)
            if empleados:
                if rut_norm in empleados:
                    reg["rut_valido"] = True
                    # Si no se capturó nombre, tomar el de la tabla
                    if not reg.get("nombre"):
                        reg["nombre"] = empleados[rut_norm]
                else:
                    reg["rut_valido"] = False
                    reg["advertencias"].append(f"⚠ RUT {rut_raw} no está en tabla_empleados")

        # Enriquecer conceptos de detalle con tabla_conceptos
        if conceptos:
            for det in reg.get("detalle", []):
                cod = det.get("concepto","").strip().lower()
                if cod in conceptos:
                    det["concepto_desc"] = conceptos[cod]
                else:
                    # Búsqueda parcial
                    matches = [v for k, v in conceptos.items() if cod and cod in k]
                    if matches:
                        det["concepto_desc"] = matches[0]


# ── Pestaña UI ─────────────────────────────────────────────────────────────────

class LiquidacionesTab(QWidget):

    # Columnas del historial (coincide con liq_historial_fetch)
    COLS_HIST = [
        ("id",               "ID"),
        ("fecha_carga",      "Fecha Carga"),
        ("origen",           "Tipo"),
        ("archivo",          "Archivo"),
        ("mes_periodo",      "Período"),
        ("rut_trabajador",   "RUT"),
        ("nombre",           "Nombre"),
        ("sueldo_base",      "Sueldo Base"),
        ("total_haberes",    "Total Haberes"),
        ("total_descuentos", "Total Desc."),
        ("liquido_pagar",    "Líquido"),
        ("afp_monto",        "AFP"),
        ("salud_monto",      "Salud"),
    ]

    # Columnas de resultados en proceso (antes de guardar)
    COLS_PROC = [
        ("archivo",          "Archivo"),
        ("origen",           "Tipo"),
        ("rut_valido",       "✓RUT"),
        ("mes_periodo",      "Período"),
        ("rut_trabajador",   "RUT"),
        ("nombre",           "Nombre"),
        ("sueldo_base",      "Sueldo Base"),
        ("total_haberes",    "Total Haberes"),
        ("total_descuentos", "Total Desc."),
        ("liquido_pagar",    "Líquido"),
        ("afp_monto",        "AFP"),
        ("salud_monto",      "Salud"),
    ]

    def __init__(self, parent=None):
        super().__init__(parent)
        self._resultados: list[dict] = []
        self._filas_libro_filtradas: list[dict] = []
        self._archivos: list[str] = []
        self._modo_proceso = False
        self._modo_libro   = False
        self._init_db()
        self._build_ui()
        self._cargar_historial()

    def _init_db(self):
        with sqlite3.connect(DB_PATH) as con:
            con.execute(SQL_LIQ_TABLA)
            con.execute(SQL_LIQ_DETALLE)
            con.commit()

    # ── UI ─────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Estilo global redondeado
        self.setStyleSheet(
            "QGroupBox { border:1px solid #e2e8f0; border-radius:10px; margin-top:6px;"
            " padding:10px; background:#f8fafc; font-size:12px; color:#475569; }"
            "QGroupBox::title { subcontrol-origin:margin; left:12px; padding:0 4px; }"
            "QPushButton { border:1px solid #cbd5e1; border-radius:8px; padding:6px 14px;"
            " background:#f1f5f9; color:#334155; font-size:12px; }"
            "QPushButton:hover { background:#e2e8f0; }"
            "QPushButton:pressed { background:#cbd5e1; }"
            "QLineEdit { border:1px solid #e2e8f0; border-radius:8px; padding:4px 10px;"
            " background:white; font-size:12px; }"
            "QComboBox { border:1px solid #e2e8f0; border-radius:8px; padding:4px 10px;"
            " background:white; font-size:12px; }"
            "QTableWidget { border:1px solid #e2e8f0; border-radius:8px;"
            " gridline-color:#f1f5f9; font-size:12px; }"
            "QHeaderView::section { background:#f8fafc; border:none;"
            " border-bottom:1px solid #e2e8f0; padding:6px; font-size:11px; color:#64748b; }"
        )

        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 12, 12, 8)
        lay.setSpacing(8)

        # ── Fila superior: carga + acciones ──
        top_row = QHBoxLayout()
        top_row.setSpacing(10)

        # Tarjeta Cargar archivos
        grp_carga = QGroupBox("Cargar archivos")
        g = QVBoxLayout(grp_carga)
        g.setSpacing(8)
        btn_row = QHBoxLayout()
        btn_pdf   = QPushButton("📂  + PDF(s)")
        btn_excel = QPushButton("📊  + Excel(s)")
        btn_pdf.setFixedHeight(34); btn_excel.setFixedHeight(34)
        btn_pdf.clicked.connect(self._sel_pdf)
        btn_excel.clicked.connect(self._sel_excel)
        btn_row.addWidget(btn_pdf); btn_row.addWidget(btn_excel)
        g.addLayout(btn_row)
        hoja_row = QHBoxLayout()
        hoja_row.addWidget(QLabel("Hoja:"))
        self.cmb_hoja = QComboBox()
        self.cmb_hoja.setPlaceholderText("(auto)")
        self.cmb_hoja.setEnabled(False)
        self.cmb_hoja.setFixedHeight(30)
        hoja_row.addWidget(self.cmb_hoja, 1)
        g.addLayout(hoja_row)
        self.lbl_archivos = QLabel("Sin archivos cargados")
        self.lbl_archivos.setStyleSheet("color:#94a3b8; font-size:11px;")
        g.addWidget(self.lbl_archivos)
        top_row.addWidget(grp_carga, 2)

        # Tarjeta Acciones
        grp_acc = QGroupBox("Acciones")
        ga = QVBoxLayout(grp_acc)
        ga.setSpacing(8)
        self.btn_procesar = QPushButton("▶  Procesar")
        self.btn_procesar.setFixedHeight(38)
        self.btn_procesar.setEnabled(False)
        self.btn_procesar.setStyleSheet(
            "QPushButton{background:#1d4ed8;color:white;font-weight:bold;"
            " border-radius:8px; border:none; font-size:13px;}"
            "QPushButton:hover{background:#1e40af;}"
            "QPushButton:disabled{background:#94a3b8;color:#e2e8f0;border:none;}")
        self.btn_procesar.clicked.connect(self._procesar)
        ga.addWidget(self.btn_procesar)
        btn_row2 = QHBoxLayout()
        btn_limpiar_sel = QPushButton("✖  Limpiar")
        btn_limpiar_sel.setFixedHeight(30)
        btn_limpiar_sel.clicked.connect(self._limpiar_sel)
        btn_row2.addWidget(btn_limpiar_sel)
        ga.addLayout(btn_row2)
        top_row.addWidget(grp_acc, 1)
        lay.addLayout(top_row)

        # ── Barra progreso ──
        self.barra = QProgressBar()
        self.barra.setTextVisible(True); self.barra.setValue(0)
        self.barra.setVisible(False); self.barra.setFixedHeight(18)
        self.barra.setStyleSheet("QProgressBar { border-radius:9px; background:#e2e8f0; }"
                                 "QProgressBar::chunk { background:#1d4ed8; border-radius:9px; }")
        lay.addWidget(self.barra)

        # ── Filtros en una linea ──
        grp_filtros = QGroupBox("Filtrar resultados")
        gf = QHBoxLayout(grp_filtros)
        gf.setSpacing(8)
        gf.addWidget(QLabel("Período:"))
        self.filt_periodo = QLineEdit(); self.filt_periodo.setPlaceholderText("ej: 2025-03")
        self.filt_periodo.setFixedWidth(100); self.filt_periodo.setFixedHeight(28)
        gf.addWidget(self.filt_periodo)
        gf.addWidget(QLabel("RUT:"))
        self.filt_rut = QLineEdit(); self.filt_rut.setPlaceholderText("ej: 12345678-9")
        self.filt_rut.setFixedWidth(120); self.filt_rut.setFixedHeight(28)
        gf.addWidget(self.filt_rut)
        gf.addWidget(QLabel("Nombre:"))
        self.filt_nombre = QLineEdit(); self.filt_nombre.setPlaceholderText("ej: García")
        self.filt_nombre.setFixedWidth(140); self.filt_nombre.setFixedHeight(28)
        gf.addWidget(self.filt_nombre)
        btn_filtrar = QPushButton("🔍  Filtrar"); btn_filtrar.setFixedHeight(28)
        btn_filtrar.clicked.connect(self._cargar_historial)
        btn_limpiar_filtros = QPushButton("✖  Limpiar"); btn_limpiar_filtros.setFixedHeight(28)
        btn_limpiar_filtros.clicked.connect(self._limpiar_filtros)
        gf.addWidget(btn_filtrar); gf.addWidget(btn_limpiar_filtros)
        gf.addStretch()
        self.filt_periodo.returnPressed.connect(self._cargar_historial)
        self.filt_rut.returnPressed.connect(self._cargar_historial)
        self.filt_nombre.returnPressed.connect(self._cargar_historial)
        lay.addWidget(grp_filtros)

        # ── Splitter: tabla principal / panel inferior ──
        splitter = QSplitter(Qt.Orientation.Vertical)

        # Tabla principal
        self.tabla = QTableWidget()
        self.tabla.setAlternatingRowColors(True)
        self.tabla.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tabla.setSortingEnabled(True)
        self.tabla.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tabla.cellClicked.connect(self._on_click_fila)
        self.tabla.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.tabla.horizontalHeader().setStretchLastSection(True)
        splitter.addWidget(self.tabla)

        # Panel inferior: detalle + advertencias + texto raw
        panel_inf = QWidget()
        lay_inf = QHBoxLayout(panel_inf)
        lay_inf.setContentsMargins(0, 0, 0, 0); lay_inf.setSpacing(6)

        # Panel detalle haberes/descuentos
        grp_det = QGroupBox("📋  Detalle Haberes / Descuentos")
        ldet = QVBoxLayout(grp_det)
        self.tbl_detalle = QTableWidget()
        self.tbl_detalle.setColumnCount(3)
        self.tbl_detalle.setHorizontalHeaderLabels(["Tipo","Concepto","Monto"])
        self.tbl_detalle.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.tbl_detalle.horizontalHeader().setStretchLastSection(True)
        self.tbl_detalle.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tbl_detalle.setAlternatingRowColors(True)
        self.tbl_detalle.setMaximumHeight(160)
        ldet.addWidget(self.tbl_detalle)
        lay_inf.addWidget(grp_det, 2)

        # Panel advertencias + texto raw
        panel_adv_raw = QWidget()
        lay_ar = QVBoxLayout(panel_adv_raw)
        lay_ar.setContentsMargins(0,0,0,0); lay_ar.setSpacing(4)

        grp_adv = QGroupBox("⚠  Advertencias")
        lv = QVBoxLayout(grp_adv)
        self.txt_adv = QTextEdit(); self.txt_adv.setReadOnly(True)
        self.txt_adv.setMaximumHeight(74); lv.addWidget(self.txt_adv)

        grp_raw = QGroupBox("🔍  Texto extraído del PDF seleccionado")
        rv = QVBoxLayout(grp_raw)
        self.txt_raw = QTextEdit(); self.txt_raw.setReadOnly(True)
        self.txt_raw.setMaximumHeight(74)
        self.txt_raw.setFont(QFont("Courier New", 8)); rv.addWidget(self.txt_raw)

        lay_ar.addWidget(grp_adv); lay_ar.addWidget(grp_raw)
        lay_inf.addWidget(panel_adv_raw, 1)

        splitter.addWidget(panel_inf)
        splitter.setSizes([380, 180])
        lay.addWidget(splitter, 1)

        # ── Barra inferior ──
        bar = QHBoxLayout()
        self.lbl_estado = QLabel("Sin datos")
        self.lbl_estado.setStyleSheet("color:#374151;font-size:11px;")

        self.btn_guardar  = QPushButton("💾  Guardar en BD")
        btn_exportar      = QPushButton("📤  Exportar Excel")
        self.btn_eliminar = QPushButton("🗑  Eliminar seleccionados")
        self.btn_eliminar.setStyleSheet(
            "QPushButton{color:#dc2626;border-color:#e2e8f0;}"
            "QPushButton:hover{background:#fef2f2;border-color:#fca5a5;}")
        btn_limpiar  = QPushButton("🗑  Limpiar vista")
        for b in (self.btn_guardar, btn_exportar, self.btn_eliminar, btn_limpiar):
            b.setFixedHeight(32)

        self.btn_guardar.clicked.connect(self._guardar)
        btn_exportar.clicked.connect(self._exportar)
        self.btn_eliminar.clicked.connect(self._eliminar_seleccionados)
        btn_limpiar.clicked.connect(self._limpiar_tabla)

        bar.addWidget(self.lbl_estado)
        bar.addStretch()
        bar.addWidget(self.btn_guardar)
        bar.addWidget(btn_exportar)
        bar.addWidget(self.btn_eliminar)
        bar.addWidget(btn_limpiar)
        lay.addLayout(bar)

        self._set_modo_historial()

    # ── Modo tabla ──────────────────────────────────────────────────────────────

    def _set_modo_proceso(self):
        """Configura columnas para resultados recién procesados."""
        self._modo_proceso = True
        self.tabla.setColumnCount(len(self.COLS_PROC))
        self.tabla.setHorizontalHeaderLabels([h for _, h in self.COLS_PROC])
        self.btn_guardar.setEnabled(True)
        self.btn_eliminar.setVisible(False)

    def _set_modo_historial(self):
        """Configura columnas para historial de BD."""
        self._modo_proceso = False
        self.tabla.setColumnCount(len(self.COLS_HIST))
        self.tabla.setHorizontalHeaderLabels([h for _, h in self.COLS_HIST])
        self.btn_guardar.setEnabled(False)
        self.btn_eliminar.setVisible(True)

    # ── Slots carga ────────────────────────────────────────────────────────────

    def _sel_pdf(self):
        rutas, _ = QFileDialog.getOpenFileNames(self, "Seleccionar PDF(s)", "", "PDF (*.pdf)")
        if rutas:
            self._archivos.extend(rutas); self._act_lbl()

    def _sel_excel(self):
        rutas, _ = QFileDialog.getOpenFileNames(
            self, "Seleccionar Excel(s)", "", "Excel (*.xlsx *.xls *.xlsm)")
        if rutas:
            self._archivos.extend(rutas)
            hojas = _ParserExcel.listar_hojas(rutas[0])
            self.cmb_hoja.clear()
            self.cmb_hoja.addItem("(auto-detectar)")
            self.cmb_hoja.addItems(hojas)
            self.cmb_hoja.setEnabled(True)
            self._act_lbl()

    def _act_lbl(self):
        n = len(self._archivos)
        nombres = ", ".join(os.path.basename(r) for r in self._archivos[:3])
        extra = f" (+{n-3} más)" if n > 3 else ""
        self.lbl_archivos.setText(f"{n} archivo(s): {nombres}{extra}")
        self.btn_procesar.setEnabled(n > 0)

    def _limpiar_sel(self):
        self._archivos.clear()
        self.lbl_archivos.setText("Sin archivos seleccionados")
        self.btn_procesar.setEnabled(False)
        self.cmb_hoja.clear(); self.cmb_hoja.setEnabled(False)

    def _procesar(self):
        hoja = self.cmb_hoja.currentText()
        if hoja in ("", "(auto-detectar)"):
            hoja = None
        self.barra.setVisible(True); self.barra.setValue(0)
        self.btn_procesar.setEnabled(False)
        self._worker = _WorkerLiq(self._archivos[:], hoja)
        self._worker.progreso.connect(lambda pct, msg: (
            self.barra.setValue(pct), self.barra.setFormat(f"{msg} — {pct}%")))
        self._worker.terminado.connect(self._on_terminado)
        self._worker.start()

    def _on_terminado(self, resultados):
        self._resultados = resultados
        self.barra.setVisible(False)
        self.btn_procesar.setEnabled(True)

        # Verificar si hay filas de libro (PDFs parseados al formato estándar)
        filas_libro = []
        for r in resultados:
            if r.get("_filas_libro"):
                filas_libro.extend(r["_filas_libro"])

        adv = []
        for r in resultados:
            for a in r.get("advertencias", []):
                adv.append(f"{r.get('archivo','')}: {a}")
            if r.get("error"):
                adv.append(f"❌ {r.get('archivo','')}: {r['error']}")
        self.txt_adv.setPlainText("\n".join(adv) if adv else "✅ Sin advertencias")

        if filas_libro:
            # ── Punto 2b: detectar conceptos sin mapeo y ofrecer agregarlos ──
            conceptos_sin_mapeo = {}
            for f in filas_libro:
                if f.get("_sin_mapeo"):
                    nombre_orig = f.get("_col_origen", "")
                    if nombre_orig and nombre_orig not in conceptos_sin_mapeo:
                        conceptos_sin_mapeo[nombre_orig] = f.get("_seccion", "")

            if conceptos_sin_mapeo:
                items_dlg = [
                    {"nombre_orig": nombre, "seccion": sec}
                    for nombre, sec in conceptos_sin_mapeo.items()
                ]
                dlg = NuevoConceptoDialog(items_dlg, self)
                if dlg.exec() == QDialog.DialogCode.Accepted:
                    agregados = getattr(dlg, "conceptos_agregados", {})
                    # Actualizar filas con los conceptos que sí se agregaron
                    for f in filas_libro:
                        if f.get("_sin_mapeo"):
                            nombre_orig = f.get("_col_origen", "")
                            if nombre_orig in agregados:
                                f["Id del concepto"] = agregados[nombre_orig]
                                f["_sin_mapeo"] = False
                # Excluir del Excel las filas que siguen sin mapeo (no agregadas o canceladas)
                filas_libro = [f for f in filas_libro if not f.get("_sin_mapeo")]

            # Punto 3: detectar empresas no encontradas en BD y preguntar
            empresas_no_bd = {}
            for f in filas_libro:
                if f.get("_empresa_no_bd") and f.get("_rut_emp"):
                    rut_e = f["_rut_emp"]
                    if rut_e not in empresas_no_bd:
                        empresas_no_bd[rut_e] = f.get("_nombre_empresa", rut_e)
            for rut_e, nom_e in empresas_no_bd.items():
                resp = QMessageBox.question(
                    self,
                    "Empresa no encontrada",
                    f"La empresa con RUT '{rut_e}' ({nom_e}) no está en la tabla de empresas.\n\n"
                    f"¿Desea agregarla automáticamente?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if resp == QMessageBox.StandardButton.Yes:
                    try:
                        with sqlite3.connect(DB_PATH) as con:
                            # id_empresa y rut_emp con el mismo valor, resto vacío
                            con.execute(
                                "INSERT OR IGNORE INTO empresas "
                                "(id_empresa,rut_emp,raz_soc,nomb_fant,region,comuna,ciudad,"
                                "direcc_emp,fono,email_repleg,afp,salud,ccaf_emp,mutual_emp,"
                                "cot_mut,cod_ae,rut_repleg,nom_repleg,email_emp) "
                                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                (rut_e, rut_e, nom_e, nom_e,
                                 "","","","","","","","","","","","","","","")
                            )
                            con.commit()
                        # Actualizar las filas con el nuevo id
                        for f in filas_libro:
                            if f.get("_rut_emp") == rut_e:
                                f["Empresa"] = rut_e
                                f["_empresa_no_bd"] = False
                    except Exception as e:
                        QMessageBox.warning(self, "Error", f"No se pudo agregar la empresa:\n{e}")
            # Mostrar en formato de 16 columnas estándar
            self._set_modo_libro()
            # DEBUG: contar aportes en filas_libro
            with open("debug_mutual_log.txt", "a", encoding="utf-8") as _f:
                _aportes = [f["Id del concepto"] for f in filas_libro if f.get("_es_aporte")]
                _f.write(f"DEBUG filas_libro: total={len(filas_libro)} aportes={_aportes}\n")
            self._filas_libro_filtradas = filas_libro  # guardar filas ya filtradas
            self._poblar_tabla_libro(filas_libro)
            sin_mapeo = len([f for f in filas_libro if f.get("_sin_mapeo")])
            self.lbl_estado.setText(
                f"✅ {len(filas_libro)} fila(s) de {len(resultados)} liquidación(es)"
                + (f" — ⚠ {sin_mapeo} concepto(s) sin mapeo" if sin_mapeo else "")
                + "  →  presiona 💾 Guardar o 📤 Exportar"
            )
        else:
            self._set_modo_proceso()
            self._poblar_tabla_proceso(resultados)
            ok  = len([r for r in resultados if r.get("origen") != "ERROR"])
            inv = len([r for r in resultados if not r.get("rut_valido", True)])
            self.lbl_estado.setText(
                f"✅ {ok} registro(s) procesados"
                + (f" — ⚠ {inv} RUT(s) no encontrado(s)" if inv else "")
                + "  →  presiona 💾 Guardar para persistir"
            )

    def _set_modo_libro(self):
        """Columnas de 16 campos estándar para PDFs parseados."""
        self._modo_proceso = True
        self._modo_libro   = True
        self.tabla.setColumnCount(len(COLS_SALIDA))
        self.tabla.setHorizontalHeaderLabels(COLS_SALIDA)
        self.btn_guardar.setEnabled(True)
        self.btn_eliminar.setVisible(False)

    def _poblar_tabla_libro(self, filas):
        """Llena la tabla con las 16 columnas estándar."""
        self.tabla.setRowCount(len(filas))
        for r, fila in enumerate(filas):
            no_mapeo = fila.get("_sin_mapeo", False)
            for c, col in enumerate(COLS_SALIDA):
                val  = fila.get(col, "")
                item = QTableWidgetItem("" if val is None else str(val))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if no_mapeo:
                    item.setBackground(QColor("#FFF3CD"))
                self.tabla.setItem(r, c, item)

    def _poblar_tabla_proceso(self, resultados):
        alias = {"rut_trabajador":"rut","mes_periodo":"periodo",
                 "liquido_pagar":"liquido","afp_monto":"afp","salud_monto":"salud"}
        self.tabla.setRowCount(len(resultados))
        for r, reg in enumerate(resultados):
            es_err = reg.get("origen") == "ERROR"
            rut_inv = reg.get("rut_valido") is False
            for c, (campo, _) in enumerate(self.COLS_PROC):
                if campo == "rut_valido":
                    val = "❌" if rut_inv else ("✅" if reg.get("rut_valido") else "—")
                else:
                    val = reg.get(campo, reg.get(alias.get(campo, campo), ""))
                item = QTableWidgetItem(self._fmt(val))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if es_err:
                    item.setBackground(QColor("#ffcccc"))
                elif rut_inv and campo in ("rut_trabajador", "rut_valido"):
                    item.setBackground(QColor("#fef9c3"))
                self.tabla.setItem(r, c, item)

    def _on_click_fila(self, row, _col):
        """Muestra detalle y texto raw del registro seleccionado."""
        if self._modo_proceso:
            if row < len(self._resultados):
                reg = self._resultados[row]
                self.txt_raw.setPlainText(reg.get("texto_crudo", "(Excel — sin texto crudo)"))
                self._poblar_detalle(reg.get("detalle", []))
                adv = reg.get("advertencias", [])
                self.txt_adv.setPlainText("\n".join(adv) if adv else "✅ Sin advertencias")
        else:
            # Historial: cargar detalle desde BD
            id_item = self.tabla.item(row, 0)
            if id_item:
                liq_id = id_item.text()
                try:
                    detalles = liq_detalle_fetch(int(liq_id))
                    self._poblar_detalle([{"tipo": d[0], "concepto": d[1], "monto": d[2]} for d in detalles])
                except Exception:
                    pass
            self.txt_raw.clear()
            self.txt_adv.clear()

    def _poblar_detalle(self, detalle):
        """Llena la tabla de detalle haberes/descuentos."""
        self.tbl_detalle.setRowCount(0)
        if not detalle:
            self.tbl_detalle.setRowCount(1)
            self.tbl_detalle.setItem(0, 0, QTableWidgetItem("(sin detalle)"))
            return
        # Agrupar por tipo para mejor visualización
        haberes    = [d for d in detalle if d.get("tipo") == "HABER"]
        descuentos = [d for d in detalle if d.get("tipo") == "DESCUENTO"]
        otros      = [d for d in detalle if d.get("tipo") not in ("HABER","DESCUENTO")]
        filas = haberes + descuentos + otros
        self.tbl_detalle.setRowCount(len(filas))
        colores = {
            "HABER_AFECTO":    "#d1fae5",
            "HABER_EXENTO":    "#e0f2fe",
            "HABER":           "#d1fae5",
            "DESCUENTO_LEGAL": "#fee2e2",
            "DESCUENTO_OTRO":  "#fef9c3",
            "DESCUENTO":       "#fee2e2",
            "APORTE":          "#ede9fe",
            "OTRO":            "#f1f5f9",
            None:              "#f8fafc",
        }
        for r, d in enumerate(filas):
            tipo    = d.get("tipo","")
            conc    = d.get("concepto_desc", d.get("concepto",""))
            monto   = d.get("monto")
            color   = QColor(colores.get(tipo, "#f8fafc"))
            for c, val in enumerate([tipo, conc, self._fmt(monto)]):
                item = QTableWidgetItem(str(val) if val is not None else "")
                item.setBackground(color)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if c == 2:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self.tbl_detalle.setItem(r, c, item)

    # ── Filtros historial ──────────────────────────────────────────────────────

    def _limpiar_filtros(self):
        self.filt_periodo.clear(); self.filt_rut.clear(); self.filt_nombre.clear()
        self._cargar_historial()

    # ── Eliminar ───────────────────────────────────────────────────────────────

    def _eliminar_seleccionados(self):
        filas_sel = list({idx.row() for idx in self.tabla.selectedIndexes()})
        if not filas_sel:
            QMessageBox.information(self, "Eliminar", "Selecciona al menos un registro."); return
        ids = []
        for row in filas_sel:
            id_item = self.tabla.item(row, 0)
            if id_item and id_item.text().isdigit():
                ids.append(int(id_item.text()))
        if not ids:
            return
        resp = QMessageBox.question(
            self, "Confirmar eliminación",
            f"¿Eliminar {len(ids)} registro(s) y su detalle permanentemente?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if resp == QMessageBox.StandardButton.Yes:
            try:
                liq_delete(ids)
                self._cargar_historial()
                self.lbl_estado.setText(f"🗑 {len(ids)} registro(s) eliminado(s)")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    # ── Guardar ────────────────────────────────────────────────────────────────

    def _guardar(self):
        if not self._resultados:
            QMessageBox.information(self, "Guardar", "No hay resultados para guardar."); return
        try:
            with sqlite3.connect(DB_PATH) as con:
                con.execute(SQL_LIQ_TABLA); con.execute(SQL_LIQ_DETALLE)
                n_ok = 0
                for r in self._resultados:
                    if r.get("origen") == "ERROR":
                        continue
                    raw = json.dumps(r.get("raw_extra", {}), ensure_ascii=False, default=str)
                    cur = con.execute("""
                        INSERT INTO liquidaciones
                            (fecha_carga,origen,archivo,mes_periodo,rut_trabajador,nombre,
                             sueldo_base,total_haberes,total_descuentos,liquido_pagar,
                             afp_monto,salud_monto,datos_raw)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""", (
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        r.get("origen",""),
                        r.get("archivo",""),
                        str(r.get("mes_periodo", r.get("periodo",""))),
                        str(r.get("rut_trabajador", r.get("rut",""))),
                        str(r.get("nombre","")),
                        self._to_f(r.get("sueldo_base")),
                        self._to_f(r.get("total_haberes")),
                        self._to_f(r.get("total_descuentos")),
                        self._to_f(r.get("liquido_pagar", r.get("liquido"))),
                        self._to_f(r.get("afp_monto", r.get("afp"))),
                        self._to_f(r.get("salud_monto", r.get("salud"))),
                        raw,
                    ))
                    liq_id = cur.lastrowid
                    for det in r.get("detalle", []):
                        con.execute(
                            "INSERT INTO liquidaciones_detalle (liquidacion_id,tipo,concepto,monto) VALUES (?,?,?,?)",
                            (liq_id, det.get("tipo"), det.get("concepto"), det.get("monto")))
                    n_ok += 1
                con.commit()
            QMessageBox.information(self, "Guardado", f"✅ {n_ok} registro(s) guardados en la BD.")
            self._resultados.clear()
            self._set_modo_historial()
            self._cargar_historial()
        except Exception as e:
            QMessageBox.critical(self, "Error al guardar", str(e))

    def _cargar_historial(self):
        """Carga (o recarga con filtros) el historial desde la BD."""
        self._set_modo_historial()
        try:
            rows = liq_historial_fetch(
                filtro_periodo=self.filt_periodo.text().strip(),
                filtro_rut=self.filt_rut.text().strip(),
                filtro_nombre=self.filt_nombre.text().strip(),
            )
            self.tabla.setRowCount(len(rows))
            for r, fila in enumerate(rows):
                for c, val in enumerate(fila):
                    item = QTableWidgetItem(self._fmt(val))
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    self.tabla.setItem(r, c, item)
            self.lbl_estado.setText(f"{len(rows)} registro(s) en historial")
            self.tbl_detalle.setRowCount(0)
            self.txt_raw.clear(); self.txt_adv.clear()
        except Exception:
            pass

    def _exportar(self):
        if self.tabla.rowCount() == 0:
            QMessageBox.information(self, "Exportar", "No hay datos para exportar."); return
        ruta, _ = QFileDialog.getSaveFileName(
            self, "Guardar Excel", "libro_liquidaciones_detalladas.xlsx", "Excel (*.xlsx)")
        if not ruta:
            return
        try:
            # Usar filas ya filtradas (sin conceptos ignorados por el usuario)
            filas_libro = getattr(self, "_filas_libro_filtradas", [])
            if not filas_libro:
                for r in self._resultados:
                    if r.get("_filas_libro"):
                        filas_libro.extend(r["_filas_libro"])

            if filas_libro:
                total, sin_mapeo = exportar_libro_remuneraciones(filas_libro, ruta)
                msg = f"✅ {total} fila(s) exportadas en formato estándar.\n\nArchivo:\n{ruta}"
                if sin_mapeo:
                    msg += f"\n\n⚠ Conceptos sin mapeo: {', '.join(sin_mapeo)}"
                resp_ab = QMessageBox.question(self, "Exportado", msg + "\n\n¿Desea abrir el archivo?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if resp_ab == QMessageBox.StandardButton.Yes:
                    import subprocess; subprocess.Popen(["start", "", ruta], shell=True)
            else:
                # Fallback: exportar vista de tabla (totales)
                from openpyxl import Workbook
                from openpyxl.styles import Font, PatternFill, Alignment
                from openpyxl.utils import get_column_letter
                wb = Workbook(); ws = wb.active; ws.title = "Liquidaciones"
                cols = self.COLS_PROC if self._modo_proceso else self.COLS_HIST
                hdrs = [h for _, h in cols]
                ws.append(hdrs)
                for c_idx in range(1, len(hdrs) + 1):
                    cell = ws.cell(1, c_idx)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="1565C0")
                    cell.alignment = Alignment(horizontal="center")
                for r in range(self.tabla.rowCount()):
                    ws.append([(self.tabla.item(r,c).text() if self.tabla.item(r,c) else "")
                               for c in range(self.tabla.columnCount())])
                for i in range(1, len(hdrs) + 1):
                    col_cells = [ws.cell(row=r, column=i) for r in range(1, ws.max_row + 1)]
                    ancho = min(max((len(str(c.value or "")) for c in col_cells), default=8) + 2, 35)
                    ws.column_dimensions[get_column_letter(i)].width = ancho
                wb.save(ruta)
                total = self.tabla.rowCount()
                msg = f"✅ {total} registro(s) exportados.\n\nArchivo:\n{ruta}"

            dlg = QMessageBox(self)
            dlg.setWindowTitle("Exportación exitosa")
            dlg.setText(msg)
            dlg.setIcon(QMessageBox.Icon.Information)
            btn_abrir = dlg.addButton("📂  Abrir", QMessageBox.ButtonRole.ActionRole)
            dlg.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            dlg.exec()
            if dlg.clickedButton() == btn_abrir:
                import subprocess
                import os; os.startfile(ruta)
        except Exception as e:
            QMessageBox.critical(self, "Error al exportar", str(e))

    def _limpiar_tabla(self):
        self.tabla.setRowCount(0)
        self._resultados.clear()
        self.txt_raw.clear(); self.txt_adv.clear()
        self.tbl_detalle.setRowCount(0)
        self._modo_libro = False
        self._set_modo_historial()
        self.lbl_estado.setText("Tabla limpia")

    @staticmethod
    def _to_f(val):
        if val is None: return None
        try:
            return float(str(val).replace(".", "").replace(",", "."))
        except Exception:
            return None

    @staticmethod
    def _fmt(val):
        if val is None or val == "": return ""
        if isinstance(val, float):
            return f"{val:,.0f}" if val == int(val) else f"{val:,.2f}"
        return str(val)
# ══════════════════════════════════════════════════════════════════════════════
# FIN MÓDULO LIQUIDACIONES
# ══════════════════════════════════════════════════════════════════════════════


# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO LIBRO REMUNERACIONES
# Lee Excel/PDF de liquidaciones masivas y produce el Excel estándar de salida
# con una fila por concepto por trabajador.
# ══════════════════════════════════════════════════════════════════════════════

# ── Columnas del Excel de origen que son conceptos (en orden) ─────────────────
# Cada entrada: (nombre_columna_origen, id_concepto_bd, tipo)
# tipo: 'haber' | 'descuento' | 'aporte' | 'info'
COLS_ORIGEN_CONCEPTOS = [
    ("Sueldo Base",                                    "sueldoBase",       "haber"),
    ("Bono de Produccion",                             "boProd",           "haber"),
    ("Comisión por ventas",                            "comision",         "haber"),
    ("Bono Desempeño",                                 "boCumpli",         "haber"),
    ("Bono Equis",                                     "boNuevo",          "haber"),
    ("Bono metas",                                     "meComer",          "haber"),
    ("Horas Extras 50%",                               "horasEx50",        "haber"),
    ("Semana Corrida",                                 "semanaCorr",       "haber"),
    ("Gratificación",                                  "gratificacion",    "haber"),
    ("Colacion",                                       "colacion",         "haber"),
    ("Movilizacion",                                   "movilizacion",     "haber"),
    ("Cargas Familiares Simples",                      "cargasSimp",       "haber"),
    ("IAS Vacaciones Legales",                         "iasVacaciones",    "haber"),
    ("Bono Sala Cuna",                                 "salaCuna",         "haber"),
    ("Cotizacion AFP",                                 "afp",              "descuento"),
    ("Cotizacion SALUD",                               "isapre",           "descuento"),
    ("Seguro de Cesantia",                             "cesEmpleado",      "descuento"),
    ("Trabajo Pesado Empleado",                        "trabajoPesaEmpl",  "descuento"),
    ("APVI Ahorro voluntario mensual",                 "apvi",             "descuento"),
    ("Impuesto",                                       "impuesto",         "descuento"),
    ("Anticipo",                                       "anticipo",         "descuento"),
    ("Creditos personales CCAF",                       "cajaCred",         "descuento"),
    ("Prestamo huelga",                                "presEmp",          "descuento"),
    ("Descuento convenio optica",                      "Convenio1",        "descuento"),
    ("Anticipo de Finiquito",                          "anticipoFiniquito","descuento"),
    ("Descuentos Varios",                              "descuentosVarios", "descuento"),
    ("Retencion Judicial",                             "retencionJudicial","descuento"),
    ("Trabajo Pesado",                                 "trabajoPesa",      "descuento"),
    ("Aporte a CCAF",                                  "cajaComp",         "aporte"),
    ("Mutual",                                         "mutual",           "aporte"),
    ("Seguro Invalidez y Sobrevivencia",               "sis",              "aporte"),
    ("Seguro de Cesantia CI",                          "cesAporteCi",      "aporte"),
    ("Seguro de Cesantia Solidario",                   "cesAporteSol",     "aporte"),
    ("Aporte AFP Empleador",                           "aporteAFPemp",     "aporte"),
    ("Aporte FAPP Compensación Expectativa de Vida",   "aporteFAPPCEV",    "aporte"),
    ("Sueldo Líquido",                                 "totalesEmpl",      "descuento"),
    ("Sueldo Liquido",                                 "totalesEmpl",      "descuento"),
    ("Total a Pagar",                                  "totalesEmpl",      "descuento"),
    ("Liquido a Pagar",                                "totalesEmpl",      "descuento"),
]

# Conceptos que necesitan 'Afecto' = total imponible
_AFECTO_TOTAL_IMP = {
    "afp","isapre","trabajoPesaEmpl","mutual","sis","aporteAFPemp","aporteFAPPCEV"
}
# Conceptos que necesitan 'Afecto' = MIN(total haberes, tope_ces_pesos)
_AFECTO_TOPE_CES  = {"cesEmpleado","cesAporteCi","cesAporteSol"}

# Columnas de identidad en el Excel de origen
_COL_EMPRESA    = "Empresa"
_COL_RUT_EMP    = "Rut empresa"
_COL_PROCESO    = "Proceso"
_COL_NOMBRE     = "Nombre"
_COL_RUT_TRAB   = "Rut"
_COL_CONTRATO   = "Contrato"
_COL_DIAS_TRAB  = "Días Trabajados"
_COL_LIQUIDO    = "Sueldo Líquido"


# ── Helpers BD ────────────────────────────────────────────────────────────────

def _lr_get_concepto_id(nombre_col):
    """Busca en tabla_conceptos el Id más parecido al nombre de la columna origen."""
    try:
        with sqlite3.connect(DB_PATH) as con:
            # Búsqueda exacta primero
            row = con.execute(
                "SELECT Id FROM tabla_conceptos WHERE LOWER(Nombre)=LOWER(?)", (nombre_col,)
            ).fetchone()
            if row:
                return row[0]
            # Búsqueda parcial
            row = con.execute(
                "SELECT Id FROM tabla_conceptos WHERE LOWER(Nombre) LIKE LOWER(?)",
                (f"%{nombre_col[:8]}%",)
            ).fetchone()
            return row[0] if row else None
    except Exception:
        return None


def _lr_build_concepto_map():
    """Construye dict {nombre_col_origen: id_concepto} usando tabla_conceptos + mapa fijo."""
    mapa = {}
    for col_origen, id_fijo, _ in COLS_ORIGEN_CONCEPTOS:
        # Intentar desde BD primero, fallback al id fijo del mapa estático
        id_bd = _lr_get_concepto_id(col_origen)
        mapa[col_origen] = id_bd if id_bd else id_fijo
    return mapa


def _lr_get_inst_id(nombre_inst, tipo):
    """Busca el id de institución en la tabla correspondiente por nombre."""
    if not nombre_inst:
        return ""

    tablas = {
        "afp":     ("inst_afp",      "id_afp",   "nombre_afp"),
        "salud":   ("inst_salud",    "id_inst",  "nombre_inst"),
        "ccaf":    ("inst_cajas",    "id_inst",  "nombre_inst"),
        "mutual":  ("inst_mutuales", "id_inst",  "nombre_inst"),
    }
    if tipo not in tablas:
        return str(nombre_inst).strip()

    tabla, col_id, col_nom = tablas[tipo]

    def norm(s):
        """Normaliza: minúsculas, sin tildes, sin palabras genéricas."""
        s = str(s).lower().strip()
        for a, b in [('á','a'),('é','e'),('í','i'),('ó','o'),('ú','u'),('ü','u'),('ñ','n')]:
            s = s.replace(a, b)
        # Quitar palabras genéricas que no ayudan a identificar
        for w in (' s.a.',' sa',' s.a',' ltda',' afp',' isapre',' fonasa'):
            s = s.replace(w, '')
        return s.strip()

    nombre_norm = norm(nombre_inst)
    if not nombre_norm:
        return str(nombre_inst).strip()

    try:
        with sqlite3.connect(DB_PATH) as con:
            rows = con.execute(
                f"SELECT {col_id}, {col_nom} FROM {tabla}"
            ).fetchall()

            # 1. Búsqueda exacta normalizada
            for row_id, row_nom in rows:
                if norm(row_nom) == nombre_norm:
                    return row_id

            # 2. Búsqueda: nombre PDF contenido en nombre BD
            for row_id, row_nom in rows:
                if nombre_norm and nombre_norm in norm(row_nom):
                    return row_id

            # 3. Búsqueda: nombre BD contenido en nombre PDF
            for row_id, row_nom in rows:
                n = norm(row_nom)
                if n and n in nombre_norm:
                    return row_id

            # 4. Búsqueda por palabras clave (primera palabra significativa)
            palabras = [p for p in nombre_norm.split() if len(p) > 2]
            for palabra in palabras:
                for row_id, row_nom in rows:
                    if palabra in norm(row_nom):
                        return row_id

            # Sin match — retornar nombre original
            return str(nombre_inst).strip()
    except Exception:
        return str(nombre_inst).strip()


def _lr_get_empresa_id(rut_empresa):
    """Retorna id_empresa desde tabla empresas por RUT."""
    try:
        rut_norm = str(rut_empresa).strip()
        with sqlite3.connect(DB_PATH) as con:
            # Buscar exacto
            row = con.execute(
                "SELECT id_empresa FROM empresas WHERE rut_emp=?", (rut_norm,)
            ).fetchone()
            if row: return row[0]
            # Buscar sin puntos
            rut_sp = rut_norm.replace(".", "")
            row = con.execute(
                "SELECT id_empresa FROM empresas WHERE REPLACE(rut_emp,'.','')=?", (rut_sp,)
            ).fetchone()
            if row: return row[0]
            # Buscar por RUT sin puntos ni guion
            rut_sg = rut_sp.replace("-", "")
            row = con.execute(
                "SELECT id_empresa FROM empresas WHERE REPLACE(REPLACE(rut_emp,'.',''),'-','')=?",
                (rut_sg,)
            ).fetchone()
            return row[0] if row else rut_empresa
    except Exception:
        return rut_empresa


def _lr_get_cot_mutual(rut_empresa):
    """Retorna cot_mut desde tabla empresas por RUT."""
    try:
        with sqlite3.connect(DB_PATH) as con:
            row = con.execute(
                "SELECT cot_mut FROM empresas WHERE rut_emp=?", (rut_empresa,)
            ).fetchone()
            return float(row[0]) if row and row[0] else 0.0
    except Exception:
        return 0.0


def _lr_get_cot_afp(id_inst_afp):
    """Retorna cot_afp desde inst_afp."""
    try:
        with sqlite3.connect(DB_PATH) as con:
            row = con.execute(
                "SELECT cot_afp FROM inst_afp WHERE LOWER(id_afp) LIKE LOWER(?)",
                (f"%{id_inst_afp}%",)
            ).fetchone()
            return float(row[0]) if row and row[0] else 0.0
    except Exception:
        return 0.0


def _lr_get_numcontrato(rut_trabajador):
    """Retorna numcontrato_emp desde tabla_empleados por RUT."""
    try:
        with sqlite3.connect(DB_PATH) as con:
            row = con.execute(
                "SELECT numcontrato_emp FROM tabla_empleados WHERE rut_emp=?",
                (rut_trabajador,)
            ).fetchone()
            if row:
                return row[0]
            # Sin puntos con guión
            rut_norm = _normalizar_rut(rut_trabajador)
            row = con.execute(
                "SELECT numcontrato_emp FROM tabla_empleados WHERE REPLACE(REPLACE(rut_emp,'.',''),'-','')=?",
                (rut_norm.replace('-',''),)
            ).fetchone()
            return row[0] if row else 1
    except Exception:
        return 1


def _lr_get_tope_ces(mes_proceso):
    """Retorna tope_ces_pesos desde tabla parámetros para el mes dado."""
    try:
        with sqlite3.connect(DB_PATH) as con:
            row = con.execute(
                "SELECT tope_ces_pesos FROM parametros WHERE mes_proc=?", (mes_proceso,)
            ).fetchone()
            return float(row[0]) if row and row[0] else 0.0
    except Exception:
        return 0.0


def _lr_get_tope_salud(mes_proceso):
    """Retorna tope_salud_pesos desde tabla parámetros para el mes dado."""
    try:
        with sqlite3.connect(DB_PATH) as con:
            row = con.execute(
                "SELECT tope_salud_pesos FROM parametros WHERE mes_proc=?", (mes_proceso,)
            ).fetchone()
            return float(row[0]) if row and row[0] else 0.0
    except Exception:
        return 0.0


def _lr_get_tope_imp(mes_proceso):
    """Retorna tope_imp_pesos_afp desde tabla parámetros para el mes dado."""
    try:
        with sqlite3.connect(DB_PATH) as con:
            row = con.execute(
                "SELECT tope_imp_pesos_afp FROM parametros WHERE mes_proc=?", (mes_proceso,)
            ).fetchone()
            return float(row[0]) if row and row[0] else 0.0
    except Exception:
        return 0.0


def _lr_get_tope_imp(mes_proceso):
    """Retorna tope_imp_pesos_afp desde tabla parámetros para el mes dado."""
    try:
        with sqlite3.connect(DB_PATH) as con:
            row = con.execute(
                "SELECT tope_imp_pesos_afp FROM parametros WHERE mes_proc=?", (mes_proceso,)
            ).fetchone()
            return float(row[0]) if row and row[0] else 0.0
    except Exception:
        return 0.0


def _lr_get_param_aportes(mes_proceso):
    """Retorna dict con tasas de aportes empleador desde tabla parámetros."""
    try:
        with sqlite3.connect(DB_PATH) as con:
            row = con.execute(
                "SELECT sis, aporte_Ccaf, aporte_afp FROM parametros WHERE mes_proc=?",
                (mes_proceso,)
            ).fetchone()
            if row:
                return {
                    "sis":        float(row[0] or 0),
                    "aporte_ccaf": float(row[1] or 0),
                    "aporte_afp":  float(row[2] or 0),
                }
    except Exception:
        pass
    return {"sis": 0.0, "aporte_ccaf": 0.0, "aporte_afp": 0.0}


def _lr_get_datos_empleado(rut_trabajador):
    """Retorna dict con datos del empleado necesarios para aportes."""
    try:
        with sqlite3.connect(DB_PATH) as con:
            # Obtener todas las columnas disponibles
            cols = [d[1] for d in con.execute("PRAGMA table_info(tabla_empleados)").fetchall()]
            needed = ["cotizacionMutu","mutual_emp","ccaf_emp","isapre_emp",
                      "tipoCont_emp","fechaIng_emp","numcontrato_emp"]
            # Solo incluir columnas que existen en la tabla
            cols_set = set(cols)
            sel_cols = [c for c in needed if c in cols_set]
            if not sel_cols:
                with open("debug_mutual_log.txt","a",encoding="utf-8") as _f:
                    _f.write(f"ERROR: ninguna columna needed en tabla_empleados. cols={cols}\n")
                return {}
            sel = ", ".join(sel_cols)
            # Buscar por RUT exacto
            row = con.execute(
                f"SELECT {sel} FROM tabla_empleados WHERE rut_emp=?",
                (rut_trabajador,)
            ).fetchone()
            if not row:
                rut_n = str(rut_trabajador).replace('.','').replace('-','').replace(' ','').upper()
                row = con.execute(
                    f"SELECT {sel} FROM tabla_empleados "
                    "WHERE REPLACE(REPLACE(UPPER(rut_emp),'.',''),'-','')=?",
                    (rut_n,)
                ).fetchone()
                with open("debug_mutual_log.txt","a",encoding="utf-8") as _f:
                    _f.write(f"  Busqueda normalizada rut={rut_n} row={row}\n")
            if row:
                return dict(zip(sel_cols, row))
            with open("debug_mutual_log.txt","a",encoding="utf-8") as _f:
                _f.write(f"  NO encontrado rut={rut_trabajador}\n")
    except Exception as _e:
        with open("debug_mutual_log.txt", "a", encoding="utf-8") as _f:
            _f.write(f"EXCEPCION _lr_get_datos_empleado({rut_trabajador}): {_e}\n")
    return {}


def _lr_calcular_antiguedad(fecha_ing, mes_proceso):
    """Calcula antigüedad en años completos desde fecha_ing hasta fin del mes de proceso."""
    from datetime import date
    import calendar
    try:
        fecha_ing = str(fecha_ing).strip()
        partes = fecha_ing.split('-')
        if len(partes) == 3:
            if len(partes[0]) == 4:   # YYYY-MM-DD
                fi = date(int(partes[0]), int(partes[1]), int(partes[2]))
            else:                      # DD-MM-YYYY
                fi = date(int(partes[2]), int(partes[1]), int(partes[0]))
        else:
            return 0
        anio, mes = int(mes_proceso[:4]), int(mes_proceso[5:])
        ultimo_dia = calendar.monthrange(anio, mes)[1]
        fp = date(anio, mes, ultimo_dia)
        anios = fp.year - fi.year
        if (fp.month, fp.day) < (fi.month, fi.day):
            anios -= 1
        return max(0, anios)
    except Exception:
        return 0


def _lr_calcular_aportes(rut_trab, mes_proceso, afecto_afp, afecto_ces,
                          id_inst_afp, id_inst_ces):
    """
    Calcula todos los aportes del empleador y retorna lista de dicts
    con el mismo formato que las filas de conceptos.
    """
    from datetime import date
    aportes = []

    emp    = _lr_get_datos_empleado(rut_trab)
    params = _lr_get_param_aportes(mes_proceso)

    def fila_aporte(id_conc, monto, afecto, id_inst, cot_jub=0):
        return {
            "Id del concepto":           id_conc,
            "Monto del concepto":        int(round(monto)),
            "Afecto":                    int(afecto),
            "Id de institución":         id_inst,
            "Cotización de jubilación":  cot_jub,
            "_es_aporte":                True,
        }

    # ── SIS ──
    tasa_sis = params.get("sis", 0)
    if tasa_sis and afecto_afp:
        monto_sis = round(tasa_sis / 100 * afecto_afp)
        aportes.append(fila_aporte("sis", monto_sis, afecto_afp, id_inst_afp))

    # ── Mutual ──
    cot_mutu = float(emp.get("cotizacionMutu") or 0)
    id_mutual = str(emp.get("mutual_emp") or "")
    with open("debug_mutual_log.txt", "a", encoding="utf-8") as _f:
        _f.write(f"DEBUG MUTUAL: rut={rut_trab} cot_mutu={cot_mutu} afecto_afp={afecto_afp} emp_keys={list(emp.keys())} cotMutu={emp.get(chr(99)+chr(111)+chr(116)+chr(105)+chr(122)+chr(97)+chr(99)+chr(105)+chr(111)+chr(110)+chr(77)+chr(117)+chr(116)+chr(117))}\n")
    if cot_mutu and afecto_afp:
        monto_mut = round(cot_mutu / 100 * afecto_afp)
        aportes.append(fila_aporte("mutual", monto_mut, afecto_afp, id_mutual, cot_mutu))

    # ── Aporte AFP Empleador y FAPP CEV (solo desde 2025-08) ──
    if mes_proceso >= "2025-08":
        if afecto_afp:
            monto_afp_emp = round(afecto_afp * 0.1 / 100)
            aportes.append(fila_aporte("aporteAFPemp", monto_afp_emp, afecto_afp, id_inst_afp))
            monto_fapp = round(afecto_afp * 0.9 / 100)
            aportes.append(fila_aporte("aporteFAPPCEV", monto_fapp, afecto_afp, id_inst_afp))

    # ── Cesantía CI y Solidario ──
    fecha_ing  = emp.get("fechaIng_emp", "")
    tipo_cont  = str(emp.get("tipoCont_emp") or "").strip().upper()
    antiguedad = _lr_calcular_antiguedad(fecha_ing, mes_proceso) if fecha_ing else 0

    if afecto_ces:
        if antiguedad < 11:
            if tipo_cont == "I":
                pct_ci  = 1.6
                pct_sol = 0.8
            else:  # "F" u otros
                pct_ci  = 2.8
                pct_sol = 0.2
            aportes.append(fila_aporte("cesAporteCi",  round(afecto_ces * pct_ci  / 100), afecto_ces, id_inst_afp))
            aportes.append(fila_aporte("cesAporteSol", round(afecto_ces * pct_sol / 100), afecto_ces, id_inst_afp))
        else:  # >= 11 años
            # cesAporteCi = 0 (no se genera)
            pct_sol = 0.8
            aportes.append(fila_aporte("cesAporteSol", round(afecto_ces * pct_sol / 100), afecto_ces, id_inst_afp))

    # ── CCAF (solo si isapre_emp = fonasa) ──
    isapre_emp = str(emp.get("isapre_emp") or "").strip().lower()
    ccaf_emp   = str(emp.get("ccaf_emp") or "")
    tasa_ccaf  = params.get("aporte_ccaf", 0)
    if "fonasa" in isapre_emp and tasa_ccaf and afecto_afp:
        monto_ccaf = round(tasa_ccaf / 100 * afecto_afp)
        aportes.append(fila_aporte("cajaComp", monto_ccaf, afecto_afp, ccaf_emp))

    return aportes
    """Normaliza RUT: sin puntos, con guión, mayúsculas."""
    r = str(rut).strip().upper().replace('.', '').replace(' ', '')
    if '-' not in r and len(r) > 1:
        r = r[:-1] + '-' + r[-1]
    return r


# ── Parser Excel de origen ────────────────────────────────────────────────────

class _LibroParser:
    """
    Lee un Excel de liquidaciones masivas (una fila por trabajador, columnas = conceptos)
    y lo transforma en lista de dicts listos para escribir al Excel de salida.
    """

    @classmethod
    def _norm_header(cls, s):
        if not s:
            return ""
        s = str(s).strip()
        # Normalizar tildes para comparación
        for a, b in [('á','a'),('é','e'),('í','i'),('ó','o'),('ú','u'),('ü','u'),('ñ','n'),
                     ('Á','A'),('É','E'),('Í','I'),('Ó','O'),('Ú','U'),('Ñ','N')]:
            s = s.replace(a, b)
        return s


    @classmethod
    def _detectar_formato_liq(cls, filas):
        """Detecta si el Excel tiene formato liquidaciones por trabajador."""
        for fila in filas[:5]:
            if fila and fila[0] and "liquidacion" in str(fila[0]).lower() and "remuneraciones" in str(fila[0]).lower():
                return True
        return False

    @classmethod
    def _parsear_periodo_liq(cls, texto):
        """Convierte 'Liquidacion de remuneraciones Enero 2025' a '2025-01'."""
        import re as _re2
        MESES = {
            "enero":"01","febrero":"02","marzo":"03","abril":"04",
            "mayo":"05","junio":"06","julio":"07","agosto":"08",
            "septiembre":"09","octubre":"10","noviembre":"11","diciembre":"12"
        }
        t = str(texto).lower()
        m = _re2.search(r"(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\s+(\d{4})", t)
        if m:
            return f"{m.group(2)}-{MESES[m.group(1)]}"
        return ""

    @classmethod
    def parsear_liq_excel(cls, ruta, hoja=None):
        """Parser para Excel de liquidaciones por trabajador (una liq por bloque)."""
        import re as _re3
        wb  = load_workbook(ruta, data_only=True)
        ws  = wb[hoja] if hoja and hoja in wb.sheetnames else wb.active
        todas_filas = list(ws.iter_rows(values_only=True))
        advertencias = []
        salida = []

        # Dividir en bloques por encabezado de liquidacion
        bloques = []
        bloque_actual = []
        for fila in todas_filas:
            if fila and fila[0] and "liquidacion" in str(fila[0]).lower() and "remuneraciones" in str(fila[0]).lower():
                if bloque_actual:
                    bloques.append(bloque_actual)
                bloque_actual = [fila]
            else:
                bloque_actual.append(fila)
        if bloque_actual:
            bloques.append(bloque_actual)

        for bloque in bloques:
            if not bloque:
                continue

            periodo = rut_emp = rut_trab = afp_nombre = isapre_nombre = ""
            afp_porc = 0.0

            for fila in bloque:
                if not fila or all(v is None for v in fila):
                    continue
                a = str(fila[0] or "").strip()
                b = str(fila[1] or "").strip() if len(fila) > 1 else ""
                c = str(fila[2] or "").strip() if len(fila) > 2 else ""

                al = a.lower()
                if "liquidacion" in al and "remuneraciones" in al:
                    periodo = cls._parsear_periodo_liq(a)
                elif al.startswith("empleador"):
                    mr = _re3.search(r"rut[:\s]+([\.\d]{6,}-[\dkK])", c, _re3.IGNORECASE)
                    if mr:
                        rut_emp = _lr_norm_rut(mr.group(1))
                elif al.startswith("trabajador"):
                    ma = _re3.search(r"AFP[:\s]+(.+?)\s+Porc[\s.:]+([\d,\.]+)", c, _re3.IGNORECASE)
                    if ma:
                        afp_nombre = ma.group(1).strip()
                        try: afp_porc = float(ma.group(2).replace(",","."))
                        except: pass
                    mi = _re3.search(r"ISAPRE[:\s]+(.+?)\s+Plan", c, _re3.IGNORECASE)
                    if mi:
                        isapre_nombre = mi.group(1).strip()
                    elif _re3.search(r"FONASA", c, _re3.IGNORECASE):
                        isapre_nombre = "Fonasa"
                elif al.startswith("rut:") or al == "rut":
                    rut_trab = _lr_norm_rut(b)

            if not rut_trab or not periodo:
                continue

            inst_afp_id = _lr_get_inst_id(afp_nombre, "afp")
            inst_sal_id = _lr_get_inst_id(isapre_nombre, "salud")
            empresa_id  = _lr_get_empresa_id(rut_emp)
            n_contrato  = _lr_get_numcontrato(rut_trab)
            try:
                with sqlite3.connect(DB_PATH) as _ce:
                    _r = _ce.execute("SELECT 1 FROM empresas WHERE rut_emp=?", (rut_emp,)).fetchone()
                    empresa_no_bd = (_r is None and bool(rut_emp))
            except:
                empresa_no_bd = False

            tope_imp = _lr_get_tope_imp(periodo) or 0.0
            tope_ces = _lr_get_tope_ces(periodo) or 0.0

            # Extraer conceptos
            _IGNORAR_A = {"empleador","trabajador","rut","c. costo","cargo",
                          "incorporaci","liquidac","tope","uf:","total","imponible",
                          "haberes","descuentos"}
            conceptos = []
            for fila in bloque:
                if not fila or all(v is None for v in fila):
                    continue
                a = str(fila[0] or "").strip()
                b = fila[1] if len(fila) > 1 else None
                c = str(fila[2] or "").strip() if len(fila) > 2 else ""
                d = fila[3] if len(fila) > 3 else None
                al = a.lower()

                if a and b and isinstance(b, (int, float)) and b > 0:
                    if not any(x in al for x in _IGNORAR_A):
                        conceptos.append((a, int(b), "HABER_AFECTO"))
                if c and d and isinstance(d, (int, float)) and d > 0:
                    cn = c.replace("\n","").strip()
                    cl = cn.lower()
                    if not any(x in cl for x in {"total","imponible","hoja","descuentos legales","tope"}):
                        conceptos.append((cn, int(d), "DESCUENTO_LEGAL"))

            if not conceptos:
                advertencias.append(f"Sin conceptos para RUT {rut_trab} periodo {periodo}")
                continue

            _NO_IMP = {"colacion","movilizacion","viatico","asignacionZona","segSaludNeto"}
            afecto_bruto = sum(v for n,v,s in conceptos
                               if s == "HABER_AFECTO"
                               and _ParserPDF._id_concepto(n)[0] not in _NO_IMP)
            afecto_imp = min(afecto_bruto, tope_imp) if tope_imp else afecto_bruto
            afecto_ces = min(afecto_bruto, tope_ces) if tope_ces else afecto_bruto
            monto_afp  = next((v for n,v,s in conceptos if _ParserPDF._id_concepto(n)[0]=="afp"), 0)
            monto_sal  = next((v for n,v,s in conceptos if _ParserPDF._id_concepto(n)[0]=="isapre"), 0)
            desc_leg   = monto_afp + monto_sal + round(afecto_ces * 0.006)

            for nombre_orig, v, seccion in conceptos:
                id_conc, sin_mapeo = _ParserPDF._id_concepto(nombre_orig)
                if id_conc in _ParserPDF._AFECTO_TOTAL_IMP:
                    afecto = int(afecto_imp)
                elif id_conc in _ParserPDF._AFECTO_TOPE_CES:
                    afecto = int(afecto_ces)
                elif id_conc == "impuesto":
                    afecto = int(afecto_imp)
                else:
                    afecto = 0

                if id_conc == "afp":
                    id_inst = inst_afp_id; cot_jub = afp_porc
                elif id_conc == "isapre":
                    id_inst = inst_sal_id; cot_jub = monto_sal
                elif id_conc in ("cesEmpleado","cesAporteCi","cesAporteSol","sis","aporteAFPemp","aporteFAPPCEV"):
                    id_inst = inst_afp_id; cot_jub = 0
                else:
                    id_inst = ""; cot_jub = 0

                salida.append({
                    "Fecha de proceso":          periodo,
                    "Id empleado":               rut_trab,
                    "Número de contrato":        n_contrato,
                    "Id del concepto":           id_conc,
                    "Monto del concepto":        v,
                    "Afecto":                    afecto,
                    "Id de institución":         id_inst,
                    "Cotización de jubilación":  cot_jub,
                    "Días de licencias":         0,
                    "Días trabajados":           30,
                    "Fecha de aplicación":       "x",
                    "Empresa":                   empresa_id,
                    "Total de rebajas por LLSS": int(desc_leg) if id_conc == "impuesto" else 0,
                    "Rentas no gravadas":        0,
                    "Rebaja por zona extrema":   0,
                    "Jornada":                   "C",
                    "_col_origen":               nombre_orig,
                    "_sin_mapeo":                sin_mapeo,
                    "_seccion":                  seccion,
                    "_rut_emp":                  rut_emp,
                    "_empresa_no_bd":            empresa_no_bd,
                })

            # Aportes empleador
            afecto_afp_val = next((f["Afecto"] for f in salida
                                   if f.get("Id del concepto")=="afp"
                                   and f.get("Id empleado")==rut_trab
                                   and f.get("Fecha de proceso")==periodo), 0)
            afecto_ces_val = next((f["Afecto"] for f in salida
                                   if f.get("Id del concepto")=="cesEmpleado"
                                   and f.get("Id empleado")==rut_trab
                                   and f.get("Fecha de proceso")==periodo), 0)
            if not afecto_ces_val and afecto_afp_val:
                afecto_ces_val = min(afecto_afp_val, int(_lr_get_tope_ces(periodo) or afecto_afp_val))

            aportes = _lr_calcular_aportes(rut_trab, periodo, afecto_afp_val, afecto_ces_val, inst_afp_id, inst_afp_id)
            for ap in aportes:
                ap.update({
                    "Fecha de proceso": periodo,
                    "Id empleado":      rut_trab,
                    "Número de contrato": n_contrato,
                    "Días de licencias": 0,
                    "Días trabajados": 30,
                    "Fecha de aplicación": "x",
                    "Empresa": empresa_id,
                    "Total de rebajas por LLSS": 0,
                    "Rentas no gravadas": 0,
                    "Rebaja por zona extrema": 0,
                    "Jornada": "C",
                    "_col_origen": ap["Id del concepto"],
                    "_sin_mapeo": False,
                    "_seccion": "APORTE",
                    "_rut_emp": rut_emp,
                    "_empresa_no_bd": empresa_no_bd,
                })
            salida.extend(aportes)

        return salida, advertencias

    @classmethod
    def parsear_csv_lre(cls, ruta):
        """
        Parser para Libro de Remuneraciones Electronico (CSV Previred/SII).
        El periodo viene en el nombre del archivo: XXXXXXXXX_AAAAMM.csv
        El RUT empresa viene en el nombre: XXXXXXXXX -> XXXXXXXX-X
        """
        import csv as _csv, os as _os, re as _re4

        advertencias = []
        salida = []

        # Periodo y RUT empresa desde nombre del archivo
        base = _os.path.splitext(_os.path.basename(ruta))[0]  # ej: 995447002_202601
        partes = base.split("_")
        rut_emp_raw = partes[0] if len(partes) >= 1 else ""
        periodo_raw = partes[-1] if len(partes) >= 2 else ""

        # RUT empresa: XXXXXXXXX -> XXXXXXXX-X
        rut_emp = (rut_emp_raw[:-1] + "-" + rut_emp_raw[-1]) if rut_emp_raw else ""

        # Periodo: AAAAMM -> AAAA-MM
        periodo = f"{periodo_raw[:4]}-{periodo_raw[4:6]}" if len(periodo_raw) == 6 else ""
        if not periodo:
            return [], ["No se pudo determinar el periodo desde el nombre del archivo."]

        # Empresa en BD
        empresa_id = _lr_get_empresa_id(rut_emp)
        try:
            with sqlite3.connect(DB_PATH) as _ce:
                _r = _ce.execute("SELECT 1 FROM empresas WHERE rut_emp=?", (rut_emp,)).fetchone()
                empresa_no_bd = (_r is None and bool(rut_emp))
        except:
            empresa_no_bd = False

        # Topes del periodo
        tope_imp = _lr_get_tope_imp(periodo) or 0.0
        tope_ces = _lr_get_tope_ces(periodo) or 0.0

        # Mapeo columna CSV -> (id_concepto, es_aporte)
        # True = aporte empleador, False = concepto normal
        MAPA_COLS = {
            "Sueldo(2101)":                                              ("sueldoBase",         False),
            "Sobresueldo(2102)":                                         ("horasEx50",           False),
            "Comisiones(2103)":                                          ("comisionMi",          False),
            "Semana corrida(2104)":                                      ("semanaCorr",          False),
            "Gratificacion(2106)":                                       ("gratificacion",       False),
            "Recargo 30% dia domingo(2107)":                             ("recargoMi",           False),
            "Aguinaldo(2110)":                                           ("AguinaldoMi",         False),
            "Bonos u otras remun. fijas mensuales(2111)":                ("BonoMi",              False),
            "Tratos(2112)":                                              ("tratoMi",             False),
            "Bonos u otras remun. variables mensuales o superiores a un mes(2113)": ("bonsupMi", False),
            "Otros ingresos no constitutivos de renta(2204)":            ("otingrMi",           False),
            "Colacion(2301)":                                            ("colacion",            False),
            "Movilizacion(2302)":                                        ("movilizacion",        False),
            "Viaticos(2303)":                                            ("viaticoMi",           False),
            "Asignacion de perdida de caja(2304)":                       ("asigPerdCajaMi",      False),
            "Asignacion de desgaste herramienta(2305)":                  ("AsigDesgHerrMi",      False),
            "Asignacion familiar legal(2311)":                           ("cargasSimp",          False),
            "Sala cuna(2308)":                                           ("salaCMi",             False),
            "Asignacion trabajo a distancia o teletrabajo(2309)":        ("AsigTeletrabajoMi",   False),
            "Indemnizacion por feriado legal(2313)":                     ("iasVacaciones",       False),
            "Indemnizacion anos de servicio(2314)":                      ("iasLegal",            False),
            "Indemnizacion sustitutiva del aviso previo(2315)":          ("iasMes",              False),
            "Cotizacion obligatoria previsional (AFP o IPS)(3141)":      ("afp",                 False),
            "Cotizacion AFC - trabajador(3151)":                         ("cesEmpleado",         False),
            "Impuesto retenido por remuneraciones(3161)":                ("impuesto",            False),
            "Credito social CCAF(3110)":                                 ("cajaCred",            False),
            "Otros descuentos autorizados y solicitados por el trabajador(3183)": ("otrdescMi",  False),
            "Cotizacion adicional trabajo pesado - trabajador(3154)":    ("trabajoPesaEmpl",     False),
            "Otros descuentos(3185)":                                    ("otrosDesctoMi",       False),
            "Pensiones de alimentos(3186)":                              ("retencionJudicial",   False),
            "Descuentos por anticipos y prestamos(3188)":                ("AnticipoPrestamoMi",  False),
            "AFC - Aporte empleador(4151)":                              ("cesAporteSol",        True),
            "Aporte empleador seguro accidentes del trabajo y Ley SANNA(4152)": ("mutual",       True),
            "Aporte adicional trabajo pesado - empleador(4154)":         ("trabajoPesa",         True),
            "Aporte empleador seguro invalidez y sobrevivencia(4155)":   ("sis",                 True),
            "Total liquido(5501)":                                       ("totalesEmpl",         False),
        }

        def norm_h(s):
            """Normalizar header para comparacion."""
            s = str(s).strip()
            for a,b in [("\xe1","a"),("\xe9","e"),("\xed","i"),("\xf3","o"),("\xfa","u"),
                        ("\xfc","u"),("\xf1","n"),("\xe3","a"),
                        ("\xc1","A"),("\xc9","E"),("\xcd","I"),("\xd3","O"),("\xda","U"),("\xd1","N")]:
                s = s.replace(a, b)
            return s

        def get_inst_by_cod(tabla, col_id, cod_lre, nombre_csv=""):
            """Busca institucion por cod_lre, con desambiguacion por nombre si hay duplicados."""
            try:
                with sqlite3.connect(DB_PATH) as con:
                    rows = con.execute(
                        f"SELECT {col_id}, nombre_{'afp' if tabla=='inst_afp' else 'inst'} "
                        f"FROM {tabla} WHERE cod_lre=?", (int(cod_lre),)
                    ).fetchall()
                    if not rows:
                        return ""
                    if len(rows) == 1:
                        return rows[0][0]
                    # Multiples: buscar por nombre
                    if nombre_csv:
                        nombre_n = norm_h(nombre_csv).lower()
                        for row in rows:
                            if norm_h(str(row[1] or "")).lower() in nombre_n or nombre_n in norm_h(str(row[1] or "")).lower():
                                return row[0]
                    return rows[0][0]
            except:
                return ""

        # Leer CSV
        try:
            for _enc in ("latin-1", "cp1252", "utf-8-sig", "utf-8"):
                try:
                    with open(ruta, "r", encoding=_enc, errors="strict") as f:
                        reader = _csv.reader(f, delimiter=";")
                        headers_raw = next(reader)
                        headers_norm = [norm_h(h) for h in headers_raw]
                        rows_data = list(reader)
                    break
                except (UnicodeDecodeError, UnicodeError):
                    continue
        except Exception as e:
            return [], [f"Error leyendo CSV: {e}"]

        # Construir indice de columnas
        idx = {norm_h(k): i for i, k in enumerate(headers_raw)}
        idx_norm = {norm_h(k): i for i, k in enumerate(headers_norm)}

        def get_col(row, nombre):
            """Obtiene valor de columna por nombre normalizado."""
            n = norm_h(nombre)
            i = idx.get(nombre) or idx_norm.get(n)
            if i is None:
                return 0
            try:
                v = str(row[i]).strip().replace(".","").replace(",",".")
                return int(float(v)) if v else 0
            except:
                return 0

        # Indices especiales
        idx_rut      = idx_norm.get(norm_h("Rut trabajador(1101)"), 0)
        idx_afp      = idx_norm.get(norm_h("AFP(1141)"))
        idx_salud    = idx_norm.get(norm_h("FONASA - ISAPRE(1143)"))
        idx_ccaf     = idx_norm.get(norm_h("CCAF(1110)"))
        idx_mutual   = idx_norm.get(norm_h("Org. administrador ley 16.744(1152)"))
        idx_dias_trab = idx_norm.get(norm_h("Nro dias trabajados en el mes(1115)"), None)
        idx_dias_lic  = idx_norm.get(norm_h("Nro dias de licencia medica en el mes(1116)"), None)
        # idx_jornada removido: jornada siempre es "C"
        idx_jornada   = None  # jornada fija en "C"

        # Indices cuotas sindicales (sumar todas)
        idx_sindicales = [i for i, h in enumerate(headers_norm) if "cuota sindical" in norm_h(h).lower()]
        idx_total_haberes = idx_norm.get(norm_h("Total haberes(5201)"))

        # Indices salud combinada (3143 + 3144)
        idx_salud_7   = idx_norm.get(norm_h("Cotizacion obligatoria salud 7%(3143)"))
        idx_salud_vol = idx_norm.get(norm_h("Cotizacion voluntaria para salud(3144)"))

        # Indices APVi (3155 + 3156)
        idx_apvi_a = idx_norm.get(norm_h("Cotizacion APVi Mod A(3155)"))
        idx_apvi_b = idx_norm.get(norm_h("Cotizacion APVi Mod B hasta UF50(3156)"))

        for row in rows_data:
            if not row or not row[0].strip():
                continue

            rut_trab = row[idx_rut].strip() if idx_rut < len(row) else ""
            if not rut_trab:
                continue

            # Instituciones
            cod_afp   = row[idx_afp].strip()   if idx_afp   and idx_afp   < len(row) else ""
            cod_salud = row[idx_salud].strip()  if idx_salud and idx_salud < len(row) else ""
            cod_ccaf  = row[idx_ccaf].strip()   if idx_ccaf  and idx_ccaf  < len(row) else ""
            cod_mut   = row[idx_mutual].strip() if idx_mutual and idx_mutual < len(row) else ""

            inst_afp_id  = get_inst_by_cod("inst_afp",    "id_afp",  cod_afp)
            inst_sal_id  = get_inst_by_cod("inst_salud",  "id_inst", cod_salud)
            inst_ccaf_id = get_inst_by_cod("inst_cajas",  "id_inst", cod_ccaf)
            inst_mut_id  = get_inst_by_cod("inst_mutuales","id_inst", cod_mut)

            n_contrato  = _lr_get_numcontrato(rut_trab)
            dias_trab   = int(row[idx_dias_trab].strip() or 30) if idx_dias_trab and idx_dias_trab < len(row) else 30
            dias_lic    = int(row[idx_dias_lic].strip()  or 0)  if idx_dias_lic  and idx_dias_lic  < len(row) else 0
            jornada     = row[idx_jornada].strip() if idx_jornada and idx_jornada < len(row) else "C"

            # Calcular totales para Afecto
            _NO_IMP = {"colacion","movilizacion","viaticoMi","asigPerdCajaMi","AsigDesgHerrMi",
                       "salaCMi","AsigTeletrabajoMi","cargasSimp","otingrMi","iasVacaciones",
                       "iasLegal","iasMes"}
            haberes_bruto = sum(
                get_col(row, h) for h in headers_raw
                if any(c in norm_h(h) for c in ["(210","(220","(230"])
            )
            afecto_bruto = sum(
                get_col(row, h) for h in headers_raw
                if "(210" in norm_h(h) or "(2106" in norm_h(h)
            )
            # Usar tope imponible directamente de columna del CSV si existe
            afecto_imp = min(haberes_bruto, tope_imp) if tope_imp else haberes_bruto
            afecto_ces = min(haberes_bruto, tope_ces) if tope_ces else haberes_bruto

            monto_afp      = get_col(row, "Cotizacion obligatoria previsional (AFP o IPS)(3141)")
            monto_sal_7    = get_col(row, "Cotizacion obligatoria salud 7%(3143)")
            monto_sal_vol  = get_col(row, "Cotizacion voluntaria para salud(3144)")
            monto_sal      = monto_sal_7 + monto_sal_vol
            monto_ces      = get_col(row, "Cotizacion AFC - trabajador(3151)")
            monto_apvi_b   = get_col(row, "Cotizacion APVi Mod B hasta UF50(3156)")
            monto_trab_pes = get_col(row, "Cotizacion adicional trabajo pesado - trabajador(3154)")
            tope_sal_pesos = _lr_get_tope_salud(periodo) or 0.0
            monto_sal_tope = int(min(monto_sal, tope_sal_pesos)) if tope_sal_pesos else monto_sal
            desc_leg       = monto_afp + monto_ces + monto_apvi_b + monto_trab_pes + monto_sal_tope

            total_haberes_csv = 0
            if idx_total_haberes is not None and idx_total_haberes < len(row):
                try: total_haberes_csv = int(str(row[idx_total_haberes]).strip().replace(".","") or 0)
                except: pass

            def make_fila(id_conc, monto_val, id_inst="", cot_jub=0, es_aporte=False):
                if id_conc in ("afp","sis","aporteAFPemp","aporteFAPPCEV","cesEmpleado","cesAporteCi","cesAporteSol"):
                    af = int(afecto_imp) if id_conc == "afp" else int(afecto_ces)
                elif id_conc == "impuesto":
                    af = int(afecto_imp)
                elif id_conc == "totalesEmpl":
                    af = total_haberes_csv
                else:
                    af = 0
                return {
                    "Fecha de proceso":          periodo,
                    "Id empleado":               rut_trab,
                    "Número de contrato":        n_contrato,
                    "Id del concepto":           id_conc,
                    "Monto del concepto":        monto_val,
                    "Afecto":                    af,
                    "Id de institución":         id_inst,
                    "Cotización de jubilación":  cot_jub,
                    "Días de licencias":         dias_lic,
                    "Días trabajados":           dias_trab,
                    "Fecha de aplicación":       "x",
                    "Empresa":                   empresa_id,
                    "Total de rebajas por LLSS": int(desc_leg) if id_conc == "impuesto" else 0,
                    "Rentas no gravadas":        0,
                    "Rebaja por zona extrema":   0,
                    "Jornada":                   "C",
                    "_col_origen":               id_conc,
                    "_sin_mapeo":                False,
                    "_seccion":                  "APORTE" if es_aporte else "HABER_AFECTO",
                    "_rut_emp":                  rut_emp,
                    "_empresa_no_bd":            empresa_no_bd,
                }

            # Conceptos simples del mapa
            for h_raw in headers_raw:
                h_n = norm_h(h_raw)
                match = MAPA_COLS.get(h_n)
                if match:
                    id_conc, es_aporte = match
                    monto_val = get_col(row, h_raw)
                    if monto_val == 0:
                        continue
                    # Institucion segun concepto
                    if id_conc == "afp":
                        inst = inst_afp_id; cot = monto_sal
                    elif id_conc in ("cesEmpleado","cesAporteSol","cesAporteCi","sis","aporteAFPemp","aporteFAPPCEV"):
                        inst = inst_afp_id; cot = 0
                    elif id_conc == "mutual":
                        inst = inst_mut_id; cot = 0
                    elif id_conc == "cajaCred":
                        inst = inst_ccaf_id; cot = 0
                    else:
                        inst = ""; cot = 0
                    salida.append(make_fila(id_conc, monto_val, inst, cot, es_aporte))

            # Salud combinada (3143 + 3144)
            monto_sal_total = 0
            if idx_salud_7   is not None and idx_salud_7   < len(row):
                try: monto_sal_total += int(str(row[idx_salud_7]).strip().replace(".","") or 0)
                except: pass
            if idx_salud_vol is not None and idx_salud_vol < len(row):
                try: monto_sal_total += int(str(row[idx_salud_vol]).strip().replace(".","") or 0)
                except: pass
            if monto_sal_total > 0:
                salida.append(make_fila("isapre", monto_sal_total, inst_sal_id, monto_sal_total))

            # APVi combinado (3155 + 3156)
            monto_apvi = 0
            if idx_apvi_a is not None and idx_apvi_a < len(row):
                try: monto_apvi += int(str(row[idx_apvi_a]).strip().replace(".","") or 0)
                except: pass
            if idx_apvi_b is not None and idx_apvi_b < len(row):
                try: monto_apvi += int(str(row[idx_apvi_b]).strip().replace(".","") or 0)
                except: pass
            if monto_apvi > 0:
                salida.append(make_fila("apvi", monto_apvi))

            # Cuotas sindicales (suma de todas)
            monto_sind = 0
            for i_s in idx_sindicales:
                if i_s < len(row):
                    try: monto_sind += int(str(row[i_s]).strip().replace(".","") or 0)
                    except: pass
            if monto_sind > 0:
                salida.append(make_fila("CuotaSindMi", monto_sind))

        return salida, advertencias


    @classmethod
    def parsear_excel(cls, ruta, hoja=None):
        """Retorna lista de filas de salida y lista de advertencias."""
        wb   = load_workbook(ruta, data_only=True)
        ws   = wb[hoja] if hoja and hoja in wb.sheetnames else wb.active
        filas = list(ws.iter_rows(values_only=True))
        if not filas:
            return [], ["Archivo vacío"]

        # Detectar fila de encabezados (primera que tenga 'Rut' o 'RUT')
        idx_hdr = 0
        for i, fila in enumerate(filas[:10]):
            vals = [str(v).strip() if v else "" for v in fila]
            if any(v.lower() in ("rut","rut empresa","proceso") for v in vals):
                idx_hdr = i
                break

        headers_raw  = [str(v).strip() if v else "" for v in filas[idx_hdr]]
        headers_norm = [cls._norm_header(h) for h in headers_raw]

        # Validar que existan columnas minimas: RUT trabajador y periodo
        _headers_lower = [h.lower() for h in headers_norm]
        _tiene_rut = any(k in _headers_lower for k in (
            "rut", "rut trabajador", "ruttrabajador", "rutworker",
            _LibroParser._norm_header(_COL_RUT_TRAB).lower()
        ))
        _tiene_proceso = any(k in _headers_lower for k in (
            "proceso", "mes proceso", "mesproceso", "periodo", "período",
            _LibroParser._norm_header(_COL_PROCESO).lower()
        ))
        if not _tiene_rut or not _tiene_proceso:
            faltantes = []
            if not _tiene_rut:     faltantes.append("RUT del trabajador")
            if not _tiene_proceso: faltantes.append("Mes de proceso")
            return [], [f"❌ El archivo no contiene la información mínima para ser procesado. "
                        f"Columna(s) faltante(s): {', '.join(faltantes)}."]

        def col(nombre):
            """Índice de columna por nombre (con y sin tildes)."""
            n = cls._norm_header(nombre)
            for i, h in enumerate(headers_norm):
                if h.lower() == n.lower():
                    return i
            # Búsqueda parcial
            for i, h in enumerate(headers_norm):
                if n.lower() in h.lower() or h.lower() in n.lower():
                    return i
            return None

        # Mapear índices de columnas de identidad
        idx_empresa   = col(_COL_EMPRESA)
        idx_rut_emp   = col(_COL_RUT_EMP)
        idx_proceso   = col(_COL_PROCESO)
        idx_rut_trab  = col(_COL_RUT_TRAB)
        idx_contrato  = col(_COL_CONTRATO)
        idx_dias_trab = col(_COL_DIAS_TRAB)

        # Mapear índices de columnas de conceptos
        concepto_idxs = []  # [(col_origen, id_concepto, tipo, idx)]
        concepto_map  = _lr_build_concepto_map()
        for col_origen, id_fijo, tipo in COLS_ORIGEN_CONCEPTOS:
            i = col(col_origen)
            if i is not None:
                concepto_idxs.append((col_origen, concepto_map.get(col_origen, id_fijo), tipo, i))

        advertencias = []
        salida = []

        for fila in filas[idx_hdr + 1:]:
            if all(v is None for v in fila):
                continue

            def val(idx):
                if idx is None or idx >= len(fila):
                    return None
                return fila[idx]

            rut_trab  = _lr_norm_rut(val(idx_rut_trab) or "")
            rut_emp   = str(val(idx_rut_emp) or "").strip()
            proceso   = str(val(idx_proceso) or "").strip()
            empresa   = _lr_get_empresa_id(rut_emp)
            n_contrato = val(idx_contrato) or _lr_get_numcontrato(rut_trab)
            dias_trab_orig = int(val(idx_dias_trab) or 30)

            if not rut_trab or rut_trab == "-":
                continue

            # Calcular totales para reglas de Afecto
            total_imp        = 0.0
            total_haberes    = 0.0
            monto_salud      = 0.0
            monto_afp        = 0.0
            monto_ces        = 0.0
            monto_trab_pes   = 0.0

            # Primera pasada: calcular totales usando secciones granulares
            for col_origen, id_conc, tipo, cidx in concepto_idxs:
                monto = float(val(cidx) or 0)
                if monto == 0:
                    continue
                # Obtener sección desde tabla_conceptos (fuente de verdad)
                tipo_bd  = _ParserPDF._tipo_concepto(id_conc)
                sec_conc = _ParserPDF._tipo_to_seccion(tipo_bd) or tipo
                if sec_conc in ("HABER_AFECTO","HABER_EXENTO","haber"):
                    total_haberes += monto
                    # Solo afectos suman al imponible
                    if sec_conc in ("HABER_AFECTO","haber") and id_conc not in (
                        "colacion","movilizacion","cargasSimp","iasVacaciones","salaCuna"
                    ):
                        total_imp += monto
                # Salud: sumar todas las variantes (mapean a isapre)
                if id_conc == "isapre":
                    monto_salud += monto
                if id_conc == "afp":
                    monto_afp = monto
                if id_conc == "cesEmpleado":
                    monto_ces = monto
                if id_conc == "trabajoPesaEmpl":
                    monto_trab_pes = monto

            tope_imp_bd = _lr_get_tope_imp(proceso)
            tope_ces    = _lr_get_tope_ces(proceso)
            tope_salud  = _lr_get_tope_salud(proceso)

            # Afecto = MIN(total_imp, tope_imp_pesos_afp)
            afecto_base = min(total_imp, tope_imp_bd) if tope_imp_bd else total_imp
            afecto_ces  = min(total_haberes, tope_ces) if tope_ces else total_haberes

            # salud_legal = solo cotización base con tope (el adicional NO entra en desc_legales)
            salud_legal  = min(monto_salud, tope_salud) if tope_salud else monto_salud
            desc_legales = monto_afp + salud_legal + monto_ces + monto_trab_pes
            # Afecto impuesto = total imponible SIN tope − desc_legales
            afecto_imp   = total_imp - desc_legales

            # Días licencia: 30 - días trabajados
            dias_lic  = 30 - dias_trab_orig
            dias_trab = 30 - dias_lic   # = dias_trab_orig

            # Institución AFP del trabajador (para cotización jubilación)
            inst_afp_trab = ""
            for col_origen, id_conc, tipo, cidx in concepto_idxs:
                if id_conc == "cot_afp" and float(val(cidx) or 0) > 0:
                    # Buscar nombre de AFP — no viene en este Excel, usar id de inst_afp
                    # Se usará el id genérico; si existe en BD se puede enriquecer
                    inst_afp_trab = _lr_get_inst_id("habitat", "afp")  # fallback
                    break

            cot_mutual = _lr_get_cot_mutual(rut_emp)

            # Segunda pasada: generar una fila por concepto
            for col_origen, id_conc, tipo, cidx in concepto_idxs:
                monto = float(val(cidx) or 0)
                if monto == 0:
                    continue

                # Sección desde tabla_conceptos
                tipo_bd  = _ParserPDF._tipo_concepto(id_conc)
                sec_conc = _ParserPDF._tipo_to_seccion(tipo_bd) or tipo

                # Afecto
                if id_conc in _AFECTO_TOTAL_IMP:
                    afecto = afecto_base
                elif id_conc in _AFECTO_TOPE_CES:
                    afecto = afecto_ces
                elif id_conc == "impuesto":
                    afecto = afecto_imp
                else:
                    afecto = 0

                # Institución
                if id_conc == "afp":
                    id_inst = _lr_get_inst_id("", "afp")
                elif id_conc == "isapre":
                    id_inst = _lr_get_inst_id("", "salud")
                elif id_conc == "cajaComp":
                    id_inst = _lr_get_inst_id("", "ccaf")
                elif id_conc == "mutual":
                    id_inst = _lr_get_inst_id("", "mutual")
                elif id_conc in ("cesEmpleado","cesAporteCi","cesAporteSol","aporteAFPemp","aporteFAPPCEV","sis"):
                    id_inst = inst_afp_trab
                else:
                    id_inst = ""

                # Cotización jubilación
                if id_conc == "afp":
                    cot_jub = _lr_get_cot_afp(str(id_inst))
                elif id_conc == "cesEmpleado":
                    cot_jub = 0.6
                elif id_conc == "isapre":
                    cot_jub = monto_salud
                elif id_conc == "mutual":
                    cot_jub = cot_mutual
                else:
                    cot_jub = 0

                # Rebajas LLSS (solo en impuesto)
                rebajas_llss = desc_legales if id_conc == "impuesto" else 0

                salida.append({
                    "Fecha de proceso":          proceso,
                    "Id empleado":               rut_trab,
                    "Número de contrato":        n_contrato,
                    "Id del concepto":           id_conc,
                    "Monto del concepto":        int(monto),
                    "Afecto":                    int(afecto) if afecto else 0,
                    "Id de institución":         id_inst,
                    "Cotización de jubilación":  cot_jub,
                    "Días de licencias":         dias_lic,
                    "Días trabajados":           dias_trab,
                    "Fecha de aplicación":       "x",
                    "Empresa":                   empresa,
                    "Total de rebajas por LLSS": int(rebajas_llss),
                    "Rentas no gravadas":        0,
                    "Rebaja por zona extrema":   0,
                    "Jornada":                   "C",
                    "_col_origen":               col_origen,
                    "_sin_mapeo":                id_conc == col_origen,
                    "_seccion":                  tipo,
                })

            # ── Aportes empleador por trabajador ──
            afecto_afp_val  = next((f["Afecto"] for f in salida
                                   if f.get("Id empleado")==rut_trab
                                   and f.get("Id del concepto")=="afp"), 0)
            afecto_ces_val  = next((f["Afecto"] for f in salida
                                   if f.get("Id empleado")==rut_trab
                                   and f.get("Id del concepto")=="cesEmpleado"), 0)
            id_inst_afp_val = next((f["Id de institución"] for f in salida
                                   if f.get("Id empleado")==rut_trab
                                   and f.get("Id del concepto")=="afp"), "")

            aportes = _lr_calcular_aportes(
                rut_trab, proceso, afecto_afp_val, afecto_ces_val,
                id_inst_afp_val, id_inst_afp_val
            )
            for ap in aportes:
                ap.update({
                    "Fecha de proceso":          proceso,
                    "Id empleado":               rut_trab,
                    "Número de contrato":        n_contrato,
                    "Días de licencias":         dias_lic,
                    "Días trabajados":           dias_trab,
                    "Fecha de aplicación":       "x",
                    "Empresa":                   empresa,
                    "Total de rebajas por LLSS": 0,
                    "Rentas no gravadas":        0,
                    "Rebaja por zona extrema":   0,
                    "Jornada":                   "C",
                    "_col_origen":               ap["Id del concepto"],
                    "_sin_mapeo":                False,
                })
            salida.extend(aportes)

        return salida, advertencias


# ── Exportar al Excel de salida ───────────────────────────────────────────────

COLS_SALIDA = [
    "Fecha de proceso", "Id empleado", "Número de contrato", "Id del concepto",
    "Monto del concepto", "Afecto", "Id de institución", "Cotización de jubilación",
    "Días de licencias", "Días trabajados", "Fecha de aplicación", "Empresa",
    "Total de rebajas por LLSS", "Rentas no gravadas", "Rebaja por zona extrema", "Jornada",
]

def exportar_libro_remuneraciones(filas, ruta_destino):
    """Escribe el Excel de salida estándar con las filas procesadas."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    HDR_COLOR = "1e3a5f"
    ALT_COLOR  = "EFF6FF"
    WARN_COLOR = "FFF3CD"   # amarillo para conceptos sin mapeo
    thin   = Side(style="thin", color="C8D4E0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = Workbook()
    ws = wb.active
    ws.title = "Liquidaciones detalladas"

    # Título
    ws.merge_cells(f"A1:{chr(64+len(COLS_SALIDA))}1")
    ws["A1"].value = "Liquidaciones detalladas"
    ws["A1"].font  = Font(bold=True, color="FFFFFF", size=13, name="Arial")
    ws["A1"].fill  = PatternFill("solid", start_color=HDR_COLOR)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    # Encabezados
    for c, h in enumerate(COLS_SALIDA, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        cell.fill      = PatternFill("solid", start_color="2563eb")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = border
    ws.row_dimensions[2].height = 30

    # Datos
    sin_mapeo = []
    for r_idx, fila in enumerate(filas):
        alt      = r_idx % 2 == 0
        no_mapeo = fila.get("_sin_mapeo", False)
        for c_idx, col in enumerate(COLS_SALIDA):
            val  = fila.get(col, "")
            cell = ws.cell(row=r_idx + 3, column=c_idx + 1, value=val)
            cell.font   = Font(name="Arial", size=10)
            cell.border = border
            cell.alignment = Alignment(
                horizontal="right" if isinstance(val, (int, float)) else "center",
                vertical="center"
            )
            if no_mapeo:
                cell.fill = PatternFill("solid", start_color=WARN_COLOR)
            elif alt:
                cell.fill = PatternFill("solid", start_color=ALT_COLOR)
        if no_mapeo:
            sin_mapeo.append(fila.get("_col_origen",""))

    # Anchos de columna — usar get_column_letter para evitar MergedCell
    from openpyxl.utils import get_column_letter
    anchos = [12,16,12,18,16,16,16,16,12,12,12,14,18,14,16,10]
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A3"
    wb.save(ruta_destino)
    return len(filas), list(set(sin_mapeo))


# ── Pestaña UI ────────────────────────────────────────────────────────────────


# ── Diálogo para agregar conceptos nuevos ─────────────────────────────────────────────

class NuevoConceptoDialog(QDialog):
    """Muestra conceptos sin mapeo y permite agregarlos a tabla_conceptos."""

    _SECCION_TIPO = {
        "HABER_AFECTO":    "Haber afecto",
        "HABER_EXENTO":    "Haber exento",
        "DESCUENTO_LEGAL": "Descuento legal",
        "DESCUENTO_OTRO":  "Descuento",
        "APORTE":          "Aporte empleador",
    }

    def __init__(self, conceptos_nuevos, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Conceptos nuevos detectados")
        self.setMinimumWidth(700)
        self.conceptos_agregados = {}
        self._conceptos_nuevos = conceptos_nuevos
        self._filas_widgets = []
        self._build_ui()

    @staticmethod
    def _generar_id(nombre):
        import unicodedata
        def quitar_tildes(s):
            return ''.join(
                c for c in unicodedata.normalize('NFD', s)
                if unicodedata.category(c) != 'Mn'
            )
        palabras = quitar_tildes(nombre.strip()).split()
        if not palabras:
            return "concepto"
        prefijo = palabras[0][:2].lower()
        resto   = ''.join(p.capitalize() for p in palabras[1:])[:7]
        return prefijo + resto

    @staticmethod
    def _id_existe(id_candidato):
        try:
            with sqlite3.connect(DB_PATH) as con:
                row = con.execute(
                    "SELECT 1 FROM tabla_conceptos WHERE LOWER(id)=LOWER(?)",
                    (id_candidato,)
                ).fetchone()
                return row is not None
        except Exception:
            return False

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setSpacing(8)

        lbl = QLabel(
            f"Se encontraron <b>{len(self._conceptos_nuevos)}</b> concepto(s) que no "
            f"están en <b>tabla_conceptos</b>.<br>"
            f"Marca los que deseas agregar y ajusta los campos si es necesario."
        )
        lbl.setWordWrap(True)
        lay.addWidget(lbl)

        hdr = QHBoxLayout()
        hdr.addSpacing(24)
        for titulo, stretch in [("Id (único)", 2), ("Nombre completo", 4), ("Tipo", 3)]:
            lh = QLabel(f"<b>{titulo}</b>")
            hdr.addWidget(lh, stretch)
        lay.addLayout(hdr)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#d0d8e4;")
        lay.addWidget(sep)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        contenedor = QWidget()
        vbox = QVBoxLayout(contenedor)
        vbox.setSpacing(4)

        for item in self._conceptos_nuevos:
            nombre_orig = item.get("nombre_orig", "")
            seccion     = item.get("seccion", "")
            tipo_suger  = self._SECCION_TIPO.get(seccion, "")
            id_suger    = self._generar_id(nombre_orig)
            id_final    = id_suger
            sufijo      = 2
            while self._id_existe(id_final):
                id_final = f"{id_suger}{sufijo}"
                sufijo  += 1

            fila      = QHBoxLayout()
            chk       = QCheckBox()
            chk.setChecked(True)
            le_id     = QLineEdit(id_final)
            le_id.setMaxLength(30)
            le_nombre = QLineEdit(nombre_orig)
            le_nombre.setMaxLength(120)
            le_tipo   = QLineEdit(tipo_suger)
            le_tipo.setMaxLength(60)

            fila.addWidget(chk, 0)
            fila.addWidget(le_id,     2)
            fila.addWidget(le_nombre, 4)
            fila.addWidget(le_tipo,   3)
            vbox.addLayout(fila)
            self._filas_widgets.append((chk, le_id, le_nombre, le_tipo))

        vbox.addStretch()
        scroll.setWidget(contenedor)
        lay.addWidget(scroll, 1)

        btn_row = QHBoxLayout()
        btn_sel_all   = QPushButton("Marcar todos")
        btn_desel_all = QPushButton("Desmarcar todos")
        btn_sel_all.setFixedHeight(28)
        btn_desel_all.setFixedHeight(28)
        btn_sel_all.clicked.connect(lambda: self._marcar_todos(True))
        btn_desel_all.clicked.connect(lambda: self._marcar_todos(False))
        btn_row.addWidget(btn_sel_all)
        btn_row.addWidget(btn_desel_all)
        btn_row.addStretch()
        lay.addLayout(btn_row)

        bb = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("Agregar seleccionados")
        bb.accepted.connect(self._aceptar)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _marcar_todos(self, estado):
        for chk, _, _, _ in self._filas_widgets:
            chk.setChecked(estado)

    def _aceptar(self):
        errores    = []
        a_insertar = []

        for chk, le_id, le_nombre, le_tipo in self._filas_widgets:
            if not chk.isChecked():
                continue
            id_val     = le_id.text().strip()
            nombre_val = le_nombre.text().strip()
            tipo_val   = le_tipo.text().strip()

            if not id_val:
                errores.append(f"El campo Id está vacío para '{nombre_val}'.")
                continue
            if self._id_existe(id_val):
                errores.append(f"El Id '{id_val}' ya existe en tabla_conceptos.")
                continue
            a_insertar.append((id_val, nombre_val, tipo_val))

        if errores:
            QMessageBox.warning(self, "Errores de validación", "\n".join(errores))
            return

        try:
            with sqlite3.connect(DB_PATH) as con:
                con.executemany(
                    "INSERT INTO tabla_conceptos (id, nombre, tipo) VALUES (?,?,?)",
                    a_insertar
                )
                con.commit()
            self.conceptos_agregados = {nombre: id_val for id_val, nombre, _ in a_insertar}
        except Exception as e:
            QMessageBox.critical(self, "Error al guardar", str(e))
            return

        self.accept()

class LibroRemuneracionesTab(QWidget):
    """Pestaña para cargar Excel/PDF de liquidaciones masivas y exportar libro estándar."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._filas: list[dict] = []
        self._archivos: list[str] = []
        self._build_ui()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 10, 10, 8)
        lay.setSpacing(8)

        # ── Panel carga ──
        grp = QGroupBox("📂  Cargar archivo de liquidaciones (Excel)")
        g   = QGridLayout(grp)
        g.setSpacing(6)

        btn_excel = QPushButton("📊  Agregar Excel")
        btn_excel.setFixedHeight(34)
        btn_excel.clicked.connect(self._sel_excel)

        lbl_hoja = QLabel("Hoja:")
        self.cmb_hoja = QComboBox()
        self.cmb_hoja.setPlaceholderText("(auto)")
        self.cmb_hoja.setEnabled(False)
        self.cmb_hoja.setFixedHeight(34)

        self.btn_procesar = QPushButton("▶  Procesar")
        self.btn_procesar.setFixedHeight(34)
        self.btn_procesar.setEnabled(False)
        self.btn_procesar.setStyleSheet(
            "QPushButton{background:#2e7d32;color:white;font-weight:bold;border-radius:4px;}"
            "QPushButton:disabled{background:#9e9e9e;color:#eee;}")
        self.btn_procesar.clicked.connect(self._procesar)

        btn_limpiar = QPushButton("✖  Limpiar")
        btn_limpiar.setFixedHeight(34)
        btn_limpiar.clicked.connect(self._limpiar)

        self.lbl_archivo = QLabel("Sin archivos seleccionados")
        self.lbl_archivo.setStyleSheet("color:#374151;font-size:11px;")

        g.addWidget(btn_excel,         0, 0)
        g.addWidget(lbl_hoja,          0, 1)
        g.addWidget(self.cmb_hoja,     0, 2)
        g.addWidget(self.btn_procesar, 0, 3)
        g.addWidget(btn_limpiar,       0, 4)
        g.addWidget(self.lbl_archivo,  1, 0, 1, 5)
        lay.addWidget(grp)

        # ── Barra progreso ──
        self.barra = QProgressBar()
        self.barra.setVisible(False); self.barra.setFixedHeight(16)
        lay.addWidget(self.barra)

        # ── Tabla previsualización ──
        self.tabla = QTableWidget()
        self.tabla.setColumnCount(len(COLS_SALIDA))
        self.tabla.setHorizontalHeaderLabels(COLS_SALIDA)
        self.tabla.setAlternatingRowColors(True)
        self.tabla.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tabla.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tabla.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.tabla.horizontalHeader().setStretchLastSection(True)
        lay.addWidget(self.tabla, 1)

        # ── Advertencias ──
        grp_adv = QGroupBox("⚠  Advertencias / Conceptos sin mapeo")
        lv = QVBoxLayout(grp_adv)
        self.txt_adv = QTextEdit(); self.txt_adv.setReadOnly(True)
        self.txt_adv.setMaximumHeight(80); lv.addWidget(self.txt_adv)
        lay.addWidget(grp_adv)

        # ── Barra inferior ──
        bar = QHBoxLayout()
        self.lbl_estado = QLabel("Sin datos")
        self.lbl_estado.setStyleSheet("color:#374151;font-size:11px;")

        btn_guardar  = QPushButton("💾  Guardar en BD")
        btn_exportar = QPushButton("🔥  Exportar Excel")
        btn_eliminar = QPushButton("🗑  Eliminar seleccionados")
        btn_limpiar  = QPushButton("🗑  Limpiar vista")
        for b in (btn_guardar, btn_exportar, btn_eliminar, btn_limpiar):
            b.setFixedHeight(32)
        btn_eliminar.setStyleSheet(
            "QPushButton{color:#b91c1c;border:1px solid #fca5a5;border-radius:6px;}"
            "QPushButton:hover{background:#fef2f2;}")
        btn_guardar.clicked.connect(self._guardar_bd)
        btn_exportar.clicked.connect(self._exportar)
        btn_eliminar.clicked.connect(self._eliminar_filas)
        btn_limpiar.clicked.connect(self._limpiar_vista)

        bar.addWidget(self.lbl_estado)
        bar.addStretch()
        bar.addWidget(btn_guardar)
        bar.addWidget(btn_exportar)
        bar.addWidget(btn_eliminar)
        bar.addWidget(btn_limpiar)
        lay.addLayout(bar)

    # ── Slots ──────────────────────────────────────────────────────────────────

    def _guardar_bd(self):
        """Guarda las filas procesadas en tabla_libro_remuneraciones."""
        if not self._filas:
            QMessageBox.information(self, "Guardar", "No hay datos para guardar."); return
        try:
            with sqlite3.connect(DB_PATH) as con:
                con.execute("""
                    CREATE TABLE IF NOT EXISTS tabla_libro_remuneraciones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        fecha_proceso TEXT, rut_empleado TEXT, num_contrato TEXT,
                        id_concepto TEXT, monto INTEGER, afecto INTEGER,
                        id_institucion TEXT, cot_jubilacion REAL,
                        dias_licencias INTEGER, dias_trabajados INTEGER,
                        fecha_aplicacion TEXT, empresa TEXT,
                        rebajas_llss INTEGER, rentas_no_grav INTEGER,
                        rebaja_zona INTEGER, jornada TEXT,
                        fecha_carga TEXT
                    )""")
                fecha_carga = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                n = 0
                for f in self._filas:
                    con.execute("""
                        INSERT INTO tabla_libro_remuneraciones
                        (fecha_proceso,rut_empleado,num_contrato,id_concepto,monto,
                         afecto,id_institucion,cot_jubilacion,dias_licencias,
                         dias_trabajados,fecha_aplicacion,empresa,rebajas_llss,
                         rentas_no_grav,rebaja_zona,jornada,fecha_carga)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (
                        f.get("Fecha de proceso",""),
                        f.get("Id empleado",""),
                        f.get("Número de contrato",""),
                        f.get("Id del concepto",""),
                        f.get("Monto del concepto",0),
                        f.get("Afecto",0),
                        f.get("Id de institución",""),
                        f.get("Cotización de jubilación",0),
                        f.get("Días de licencias",0),
                        f.get("Días trabajados",30),
                        f.get("Fecha de aplicación","x"),
                        f.get("Empresa",""),
                        f.get("Total de rebajas por LLSS",0),
                        f.get("Rentas no gravadas",0),
                        f.get("Rebaja por zona extrema",0),
                        f.get("Jornada","C"),
                        fecha_carga,
                    ))
                    n += 1
                con.commit()
            QMessageBox.information(self, "Guardado", f"✅ {n} fila(s) guardadas en la BD.")
        except Exception as e:
            QMessageBox.critical(self, "Error al guardar", str(e))

    def _eliminar_filas(self):
        """Elimina las filas seleccionadas de la tabla visual."""
        filas_sel = sorted(
            {idx.row() for idx in self.tabla.selectedIndexes()}, reverse=True)
        if not filas_sel:
            QMessageBox.information(self, "Eliminar", "Selecciona al menos una fila."); return
        resp = QMessageBox.question(
            self, "Confirmar", f"¿Eliminar {len(filas_sel)} fila(s) de la vista?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if resp == QMessageBox.StandardButton.Yes:
            for row in filas_sel:
                self.tabla.removeRow(row)
                if row < len(self._filas):
                    self._filas.pop(row)
            self.lbl_estado.setText(f"{self.tabla.rowCount()} fila(s) generadas")

    def _limpiar_vista(self):
        """Limpia la tabla y reinicia el estado."""
        self._filas.clear()
        self._filas_libro_filtradas = []
        self.tabla.setRowCount(0)
        self.txt_adv.clear()
        self.lbl_estado.setText("Sin datos")
        self._archivos.clear()
        self.lbl_archivo.setText("Sin archivos seleccionados")
        self.btn_procesar.setEnabled(False)
        self.cmb_hoja.clear(); self.cmb_hoja.setEnabled(False)

    def _sel_excel(self):
        rutas, _ = QFileDialog.getOpenFileNames(
            self, "Seleccionar Excel", "", "Excel (*.xlsx *.xls *.xlsm)")
        if rutas:
            self._archivos = rutas
            self.lbl_archivo.setText(", ".join(os.path.basename(r) for r in rutas))
            wb = load_workbook(rutas[0], read_only=True)
            self.cmb_hoja.clear()
            self.cmb_hoja.addItem("(auto-detectar)")
            self.cmb_hoja.addItems(wb.sheetnames)
            self.cmb_hoja.setEnabled(True)
            self.btn_procesar.setEnabled(True)

    def _sel_pdf(self):
        rutas, _ = QFileDialog.getOpenFileNames(
            self, "Seleccionar PDF", "", "PDF (*.pdf)")
        if rutas:
            self._archivos = rutas
            self.lbl_archivo.setText(", ".join(os.path.basename(r) for r in rutas))
            self.cmb_hoja.setEnabled(False)
            self.btn_procesar.setEnabled(True)

    def _limpiar(self):
        self._archivos.clear()
        self._filas.clear()
        self.tabla.setRowCount(0)
        self.txt_adv.clear()
        self.lbl_archivo.setText("Sin archivos seleccionados")
        self.btn_procesar.setEnabled(False)
        self.cmb_hoja.clear(); self.cmb_hoja.setEnabled(False)
        self.lbl_estado.setText("Sin datos")

    def _procesar(self):
        if not self._archivos:
            return
        hoja = self.cmb_hoja.currentText()
        if hoja in ("", "(auto-detectar)"):
            hoja = None

        self.barra.setVisible(True); self.barra.setValue(10)
        self.btn_procesar.setEnabled(False)
        self._filas.clear()
        todas_adv = []

        for ruta in self._archivos:
            ext = os.path.splitext(ruta)[1].lower()
            try:
                if ext in (".xlsx", ".xls", ".xlsm"):
                    # Detectar formato: libro clasico o liquidaciones por trabajador
                    _wb_tmp = load_workbook(ruta, data_only=True)
                    _ws_tmp = _wb_tmp[hoja] if hoja and hoja in _wb_tmp.sheetnames else _wb_tmp.active
                    _filas_tmp = list(_ws_tmp.iter_rows(values_only=True))[:5]
                    if _LibroParser._detectar_formato_liq(_filas_tmp):
                        filas, adv = _LibroParser.parsear_liq_excel(ruta, hoja)
                    else:
                        filas, adv = _LibroParser.parsear_excel(ruta, hoja)
                    # Mostrar dialogo si hay error critico (archivo sin info minima)
                    for msg in adv:
                        if msg.startswith("❌"):
                            QMessageBox.warning(self, "Archivo no válido",
                                msg.replace("❌ ", ""))
                            self.barra.setVisible(False)
                            self.btn_procesar.setEnabled(True)
                            return
                else:
                    filas, adv = [], [f"⚠ PDF aún no soportado en este módulo — usa el módulo Liquidaciones"]
                self._filas.extend(filas)
                todas_adv.extend(adv)
            except Exception as e:
                todas_adv.append(f"❌ {os.path.basename(ruta)}: {e}")

        self.barra.setValue(80)
        self._poblar_tabla()
        self.barra.setValue(100); self.barra.setVisible(False)
        self.btn_procesar.setEnabled(True)

        # ── Detectar conceptos sin mapeo y ofrecer agregarlos
        conceptos_sin_mapeo_lr = {}
        for f in self._filas:
            if f.get("_sin_mapeo"):
                nom = f.get("_col_origen", "")
                if nom and nom not in conceptos_sin_mapeo_lr:
                    conceptos_sin_mapeo_lr[nom] = f.get("_seccion", "")

        if conceptos_sin_mapeo_lr:
            items_dlg = [{"nombre_orig": n, "seccion": s} for n, s in conceptos_sin_mapeo_lr.items()]
            dlg_conc = NuevoConceptoDialog(items_dlg, self)
            if dlg_conc.exec() == QDialog.DialogCode.Accepted:
                agregados_lr = getattr(dlg_conc, "conceptos_agregados", {})
                for f in self._filas:
                    if f.get("_sin_mapeo"):
                        nom = f.get("_col_origen", "")
                        if nom in agregados_lr:
                            f["Id del concepto"] = agregados_lr[nom]
                            f["_sin_mapeo"] = False
            self._filas = [f for f in self._filas if not f.get("_sin_mapeo")]
            self._poblar_tabla()

        sin_mapeo = [f.get("_col_origen","") for f in self._filas if f.get("_sin_mapeo")]
        if sin_mapeo:
            todas_adv.append(f"⚠ Conceptos sin mapeo en tabla_conceptos (se usó nombre origen): "
                             f"{', '.join(set(sin_mapeo))}")
        self.txt_adv.setPlainText("\n".join(todas_adv) if todas_adv else "✅ Sin advertencias")
        self.lbl_estado.setText(
            f"✅ {len(self._filas)} fila(s) generadas — presiona 📤 Exportar para guardar")

    def _poblar_tabla(self):
        self.tabla.setRowCount(len(self._filas))
        for r, fila in enumerate(self._filas):
            no_mapeo = fila.get("_sin_mapeo", False)
            for c, col in enumerate(COLS_SALIDA):
                val  = fila.get(col, "")
                item = QTableWidgetItem("" if val is None else str(val))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if no_mapeo:
                    item.setBackground(QColor("#FFF3CD"))
                self.tabla.setItem(r, c, item)

    def _exportar(self):
        if not self._filas:
            QMessageBox.information(self, "Exportar", "No hay datos para exportar."); return
        ruta, _ = QFileDialog.getSaveFileName(
            self, "Guardar Excel", "libro_liquidaciones_detalladas.xlsx", "Excel (*.xlsx)")
        if not ruta:
            return
        try:
            total, sin_mapeo = exportar_libro_remuneraciones(self._filas, ruta)
            msg = f"✅ {total} fila(s) exportadas correctamente.\n\nArchivo:\n{ruta}"
            if sin_mapeo:
                msg += f"\n\n⚠ Conceptos sin mapeo ({len(sin_mapeo)}):\n" + "\n".join(sin_mapeo)
            dlg = QMessageBox(self)
            dlg.setWindowTitle("Exportación exitosa")
            dlg.setText(msg)
            dlg.setIcon(QMessageBox.Icon.Information)
            btn_abrir = dlg.addButton("📂  Abrir", QMessageBox.ButtonRole.ActionRole)
            dlg.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            dlg.exec()
            if dlg.clickedButton() == btn_abrir:
                import subprocess
                import os; os.startfile(ruta)
            self.lbl_estado.setText(f"✅ Exportado: {total} fila(s)")
        except Exception as e:
            QMessageBox.critical(self, "Error al exportar", str(e))


# ══════════════════════════════════════════════════════════════════════════════
# FIN MÓDULO LIBRO REMUNERACIONES
# ══════════════════════════════════════════════════════════════════════════════


class LibroLRETab(QWidget):
    """Pestana para cargar CSV del Libro de Remuneraciones Electronico."""

    def __init__(self):
        super().__init__()
        self._archivos = []
        self._filas: list[dict] = []
        self._filas_filtradas: list[dict] = []
        self._build_ui()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 10, 10, 8); lay.setSpacing(6)

        grp = QGroupBox("Cargar archivo CSV (Libro de Remuneraciones Electronico)")
        g   = QGridLayout(grp); g.setSpacing(6)

        btn_csv = QPushButton("📋  Agregar CSV")
        btn_csv.setFixedHeight(34)
        btn_csv.clicked.connect(self._sel_csv)

        self.btn_procesar = QPushButton("▶  Procesar")
        self.btn_procesar.setFixedHeight(34)
        self.btn_procesar.setEnabled(False)
        self.btn_procesar.setStyleSheet(
            "QPushButton{background:#2e7d32;color:white;font-weight:bold;border-radius:4px;}"
            "QPushButton:disabled{background:#9e9e9e;color:#eee;}")
        self.btn_procesar.clicked.connect(self._procesar)

        btn_limpiar = QPushButton("✖  Limpiar")
        btn_limpiar.setFixedHeight(34)
        btn_limpiar.clicked.connect(self._limpiar)

        self.lbl_archivo = QLabel("Sin archivos seleccionados")
        self.lbl_archivo.setStyleSheet("color:#374151;font-size:11px;")

        g.addWidget(btn_csv,           0, 0)
        g.addWidget(self.btn_procesar, 0, 1)
        g.addWidget(btn_limpiar,       0, 2)
        g.addWidget(self.lbl_archivo,  1, 0, 1, 3)
        lay.addWidget(grp)

        self.barra = QProgressBar()
        self.barra.setTextVisible(True); self.barra.setValue(0)
        self.barra.setVisible(False); self.barra.setFixedHeight(18)
        lay.addWidget(self.barra)

        # Tabla de resultados
        self.tabla = QTableWidget(0, len(COLS_SALIDA))
        self.tabla.setHorizontalHeaderLabels(COLS_SALIDA)
        self.tabla.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.tabla.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tabla.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tabla.setAlternatingRowColors(True)
        lay.addWidget(self.tabla, 1)

        # Panel advertencias
        grp_adv = QGroupBox("⚠  Advertencias")
        ga = QVBoxLayout(grp_adv)
        self.txt_adv = QTextEdit(); self.txt_adv.setReadOnly(True)
        self.txt_adv.setFixedHeight(70)
        ga.addWidget(self.txt_adv)
        lay.addWidget(grp_adv)

        # Barra inferior
        bar = QHBoxLayout()
        self.lbl_estado = QLabel("Sin datos")
        self.lbl_estado.setStyleSheet("color:#374151;font-size:11px;")

        btn_guardar  = QPushButton("💾  Guardar en BD")
        btn_exportar = QPushButton("🔥  Exportar Excel")
        btn_eliminar = QPushButton("🗑  Eliminar seleccionados")
        btn_limpiar2 = QPushButton("🗑  Limpiar vista")
        for b in (btn_guardar, btn_exportar, btn_eliminar, btn_limpiar2):
            b.setFixedHeight(32)
        btn_eliminar.setStyleSheet(
            "QPushButton{color:#b91c1c;border:1px solid #fca5a5;border-radius:6px;}"
            "QPushButton:hover{background:#fef2f2;}")
        btn_guardar.clicked.connect(self._guardar_bd)
        btn_exportar.clicked.connect(self._exportar)
        btn_eliminar.clicked.connect(self._eliminar_filas)
        btn_limpiar2.clicked.connect(self._limpiar)

        bar.addWidget(self.lbl_estado); bar.addStretch()
        bar.addWidget(btn_guardar); bar.addWidget(btn_exportar)
        bar.addWidget(btn_eliminar); bar.addWidget(btn_limpiar2)
        lay.addLayout(bar)

    def _sel_csv(self):
        rutas, _ = QFileDialog.getOpenFileNames(
            self, "Seleccionar CSV LRE", "", "CSV (*.csv *.txt)")
        if rutas:
            self._archivos = rutas
            self.lbl_archivo.setText(", ".join(os.path.basename(r) for r in rutas))
            self.btn_procesar.setEnabled(True)

    def _procesar(self):
        if not self._archivos:
            return
        self.barra.setVisible(True); self.barra.setValue(10)
        self.btn_procesar.setEnabled(False)
        self._filas.clear()
        todas_adv = []

        for ruta in self._archivos:
            try:
                filas, adv = _LibroParser.parsear_csv_lre(ruta)
                # Detectar conceptos sin mapeo
                conceptos_sin_mapeo = {}
                for f in filas:
                    if f.get("_sin_mapeo"):
                        nom = f.get("_col_origen","")
                        if nom and nom not in conceptos_sin_mapeo:
                            conceptos_sin_mapeo[nom] = f.get("_seccion","")
                if conceptos_sin_mapeo:
                    items_dlg = [{"nombre_orig": n, "seccion": s} for n, s in conceptos_sin_mapeo.items()]
                    dlg = NuevoConceptoDialog(items_dlg, self)
                    if dlg.exec() == QDialog.DialogCode.Accepted:
                        agregados = getattr(dlg, "conceptos_agregados", {})
                        for f in filas:
                            if f.get("_sin_mapeo") and f.get("_col_origen","") in agregados:
                                f["Id del concepto"] = agregados[f["_col_origen"]]
                                f["_sin_mapeo"] = False
                    filas = [f for f in filas if not f.get("_sin_mapeo")]
                self._filas.extend(filas)
                todas_adv.extend(adv)
            except Exception as e:
                todas_adv.append(f"✖ Error en {os.path.basename(ruta)}: {e}")

        self.barra.setValue(80)
        self._filas_filtradas = list(self._filas)
        self._poblar_tabla()
        self.txt_adv.setPlainText("\n".join(todas_adv) if todas_adv else "✅ Sin advertencias")
        self.lbl_estado.setText(f"{len(self._filas)} fila(s) generadas")
        self.barra.setValue(100); self.barra.setVisible(False)
        self.btn_procesar.setEnabled(True)

    def _poblar_tabla(self):
        self.tabla.setRowCount(len(self._filas_filtradas))
        for r, fila in enumerate(self._filas_filtradas):
            for c, col in enumerate(COLS_SALIDA):
                v = fila.get(col, "")
                self.tabla.setItem(r, c, QTableWidgetItem(str(v) if v != 0 else "0"))

    def _exportar(self):
        if not self._filas:
            QMessageBox.information(self, "Exportar", "No hay datos para exportar."); return
        ruta, _ = QFileDialog.getSaveFileName(
            self, "Guardar Excel", "libro_lre_detallado.xlsx",
            "Excel (*.xlsx)")
        if not ruta:
            return
        try:
            total, sin_mapeo = exportar_libro_remuneraciones(self._filas, ruta)
            msg_exp = "✅ " + str(total) + " fila(s) exportadas."
            if sin_mapeo: msg_exp += " | ⚠ " + str(sin_mapeo) + " sin mapeo."
            resp_ab = QMessageBox.question(self, "Exportado",
                msg_exp + "\n\n¿Desea abrir el archivo?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if resp_ab == QMessageBox.StandardButton.Yes:
                import subprocess; subprocess.Popen(["start", "", ruta], shell=True)
        except Exception as e:
            QMessageBox.critical(self, "Error al exportar", str(e))

    def _guardar_bd(self):
        if not self._filas:
            QMessageBox.information(self, "Guardar", "No hay datos para guardar."); return
        try:
            with sqlite3.connect(DB_PATH) as con:
                con.execute("""
                    CREATE TABLE IF NOT EXISTS tabla_libro_remuneraciones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        fecha_proceso TEXT, rut_empleado TEXT, num_contrato TEXT,
                        id_concepto TEXT, monto INTEGER, afecto INTEGER,
                        id_institucion TEXT, cot_jubilacion REAL,
                        dias_licencias INTEGER, dias_trabajados INTEGER,
                        fecha_aplicacion TEXT, empresa TEXT,
                        rebajas_llss INTEGER, rentas_no_grav INTEGER,
                        rebaja_zona INTEGER, jornada TEXT, fecha_carga TEXT
                    )""")
                fecha_carga = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                n = 0
                for f in self._filas:
                    con.execute("""
                        INSERT INTO tabla_libro_remuneraciones
                        (fecha_proceso,rut_empleado,num_contrato,id_concepto,monto,
                         afecto,id_institucion,cot_jubilacion,dias_licencias,
                         dias_trabajados,fecha_aplicacion,empresa,rebajas_llss,
                         rentas_no_grav,rebaja_zona,jornada,fecha_carga)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (
                        f.get("Fecha de proceso",""), f.get("Id empleado",""),
                        f.get("Número de contrato",""), f.get("Id del concepto",""),
                        f.get("Monto del concepto",0), f.get("Afecto",0),
                        f.get("Id de institución",""), f.get("Cotización de jubilación",0),
                        f.get("Días de licencias",0), f.get("Días trabajados",30),
                        f.get("Fecha de aplicación","x"), f.get("Empresa",""),
                        f.get("Total de rebajas por LLSS",0), f.get("Rentas no gravadas",0),
                        f.get("Rebaja por zona extrema",0), f.get("Jornada","C"), fecha_carga,
                    ))
                    n += 1
                con.commit()
            QMessageBox.information(self, "Guardado", f"✅ {n} fila(s) guardadas.")
        except Exception as e:
            QMessageBox.critical(self, "Error al guardar", str(e))

    def _eliminar_filas(self):
        filas_sel = sorted({idx.row() for idx in self.tabla.selectedIndexes()}, reverse=True)
        if not filas_sel:
            QMessageBox.information(self, "Eliminar", "Selecciona al menos una fila."); return
        resp = QMessageBox.question(self, "Confirmar",
            f"¿Eliminar {len(filas_sel)} fila(s)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if resp == QMessageBox.StandardButton.Yes:
            for row in filas_sel:
                self.tabla.removeRow(row)
                if row < len(self._filas):
                    self._filas.pop(row)
            self.lbl_estado.setText(f"{self.tabla.rowCount()} fila(s)")

    def _limpiar(self):
        self._filas.clear(); self._filas_filtradas.clear()
        self.tabla.setRowCount(0); self.txt_adv.clear()
        self.lbl_estado.setText("Sin datos")
        self._archivos.clear()
        self.lbl_archivo.setText("Sin archivos seleccionados")
        self.btn_procesar.setEnabled(False)



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("📊 Carga de Liquidaciones en Detalle")
        self.setMinimumSize(1300, 720)
        self._build_ui()
        self._load_parametros()
        self._load_afp()
        self._load_salud()
        self._load_empresas()
        self._load_cajas()
        self._load_mutuales()
        self._load_apv()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0); layout.setSpacing(0)

        # Cabecera
        hdr_widget = QWidget()
        hdr_widget.setStyleSheet("background:#f8fafc; border-bottom:1px solid #e2e8f0;")
        hdr_widget.setFixedHeight(52)
        hdr_lay = QHBoxLayout(hdr_widget)
        hdr_lay.setContentsMargins(20, 0, 20, 0)
        lbl = QLabel("Carga de Liquidaciones en Detalle")
        lbl.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        lbl.setStyleSheet("color:#1e293b;")
        ver = QLabel("v2.0")
        ver.setStyleSheet("color:#94a3b8; font-size:12px;")
        hdr_lay.addWidget(lbl); hdr_lay.addStretch(); hdr_lay.addWidget(ver)
        layout.addWidget(hdr_widget)

        # Tabs principales
        self.tabs_main = QTabWidget()
        self.tabs_main.setStyleSheet(
            "QTabWidget::pane { border:none; background:white; margin-top:0px; }"
            "QTabBar { background:#f8fafc; }"
            "QTabBar::tab { padding:11px 36px; font-size:13px; font-weight:500;"
            " background:#f8fafc; border:none; color:#64748b;"
            " border-bottom:3px solid transparent; min-width:160px; }"
            "QTabBar::tab:selected { color:#1e40af; background:white;"
            " border-bottom:3px solid #1e40af; }"
            "QTabBar::tab:last { font-size:12px; min-width:100px; color:#94a3b8; }"
            "QTabBar::tab:hover { background:#f1f5f9; }"
        )
        layout.addWidget(self.tabs_main)

        # Panel Liquidaciones
        self.panel_liq = QWidget()
        self.panel_liq.setStyleSheet("background:white;")
        self.tabs_main.addTab(self.panel_liq, "📄  Liquidaciones")

        # Panel Libro Remuneraciones
        self.panel_libro_main = QWidget()
        self.panel_libro_main.setStyleSheet("background:white;")
        self.tabs_main.addTab(self.panel_libro_main, "  Libro Remuneraciones")

        # Panel Mantenedores
        self.panel_mant = QWidget()
        self.panel_mant.setStyleSheet("background:white;")
        self.tabs_main.addTab(self.panel_mant, "  Mantenedores ›")

        # Mantenedores: subtabs discretos
        mant_layout = QVBoxLayout(self.panel_mant)
        mant_layout.setContentsMargins(0, 0, 0, 0); mant_layout.setSpacing(0)
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet(
            "QTabWidget::pane { border:1px solid #e2e8f0; border-radius:0 8px 8px 8px; background:white; }"
            "QTabBar::tab { padding:7px 16px; font-size:12px; background:#f1f5f9;"
            " border:none; color:#94a3b8; border-bottom:2px solid transparent; }"
            "QTabBar::tab:selected { color:#475569; background:white;"
            " border-bottom:2px solid #475569; }"
            "QTabBar::tab:hover { background:#e2e8f0; }"
        )
        mant_layout.addWidget(self.tabs)

        # ── Pestaña Parámetros Mensuales ──
        self.tab_param = TableWidget(PARAM_HEADERS, "Ej: 2024-06")
        self.tab_param.btn_add.clicked.connect(self._param_add)
        self.tab_param.btn_edit.clicked.connect(self._param_edit)
        self.tab_param.btn_delete.clicked.connect(self._param_delete)
        self.tab_param.btn_refresh.clicked.connect(self._load_parametros)
        self.tab_param.search_box.textChanged.connect(self._load_parametros)
        self.tab_param.table.doubleClicked.connect(self._param_edit)
        # Botón cargar PDF
        self.btn_pdf = QPushButton("📄  Cargar PDF")
        self.btn_pdf.setObjectName("btn_pdf")
        self.btn_pdf.setFixedHeight(36)
        self.btn_pdf.clicked.connect(self._param_cargar_pdf)
        self.tab_param.layout().itemAt(0).layout().insertWidget(3, self.btn_pdf)
        self.tabs.addTab(self.tab_param, "📅  Parámetros Mensuales")

        # ── Pestaña AFP ──
        self.tab_afp = TableWidget(AFP_HEADERS, "Ej: Capital")
        self.tab_afp.btn_add.clicked.connect(self._afp_add)
        self.tab_afp.btn_edit.clicked.connect(self._afp_edit)
        self.tab_afp.btn_delete.clicked.connect(self._afp_delete)
        self.tab_afp.btn_refresh.clicked.connect(self._load_afp)
        self.tab_afp.search_box.textChanged.connect(self._load_afp)
        self.tab_afp.table.doubleClicked.connect(self._afp_edit)
        self.tabs.addTab(self.tab_afp, "🏦  Instituciones AFP")
        self.tab_salud = TableWidget(SALUD_HEADERS, "Ej: Banmedica")
        self.tab_salud.btn_add.clicked.connect(self._salud_add)
        self.tab_salud.btn_edit.clicked.connect(self._salud_edit)
        self.tab_salud.btn_delete.clicked.connect(self._salud_delete)
        self.tab_salud.btn_refresh.clicked.connect(self._load_salud)
        self.tab_salud.search_box.textChanged.connect(self._load_salud)
        self.tab_salud.table.doubleClicked.connect(self._salud_edit)
        self.tab_empresas = TableWidget(EMPRESAS_HEADERS, 'Ej: Mi Empresa')
        self.tab_empresas.btn_add.clicked.connect(self._empresas_add)
        self.tab_empresas.btn_edit.clicked.connect(self._empresas_edit)
        self.tab_empresas.btn_delete.clicked.connect(self._empresas_delete)
        self.tab_empresas.btn_refresh.clicked.connect(self._load_empresas)
        self.tab_empresas.search_box.textChanged.connect(self._load_empresas)
        self.tabs.addTab(self.tab_empresas, '  Empresas')
        self.tab_cajas = TableWidget(CAJAS_HEADERS, "Ej: Los Andes")
        self.tab_cajas.btn_add.clicked.connect(self._cajas_add)
        self.tab_cajas.btn_edit.clicked.connect(self._cajas_edit)
        self.tab_cajas.btn_delete.clicked.connect(self._cajas_delete)
        self.tab_cajas.btn_refresh.clicked.connect(self._load_cajas)
        self.tab_cajas.search_box.textChanged.connect(self._load_cajas)
        self.tabs.addTab(self.tab_cajas, "  Cajas Compensacion")
        self.tab_mutuales = TableWidget(MUTUALES_HEADERS, "Ej: ACHS")
        self.tab_mutuales.btn_add.clicked.connect(self._mutuales_add)
        self.tab_mutuales.btn_edit.clicked.connect(self._mutuales_edit)
        self.tab_mutuales.btn_delete.clicked.connect(self._mutuales_delete)
        self.tab_mutuales.btn_refresh.clicked.connect(self._load_mutuales)
        self.tab_mutuales.search_box.textChanged.connect(self._load_mutuales)
        self.tabs.addTab(self.tab_mutuales, "  Mutuales")
        self.tab_apv = TableWidget(APV_HEADERS)
        self.tab_apv.btn_add.clicked.connect(self._apv_add)
        self.tab_apv.btn_edit.clicked.connect(self._apv_edit)
        self.tab_apv.btn_delete.clicked.connect(self._apv_delete)
        self.tab_apv.search_box.textChanged.connect(self._load_apv)
        self.tab_apv.btn_refresh.clicked.connect(self._load_apv)
        self.tab_apv.search_box.textChanged.connect(self._load_apv)
        self.tabs.addTab(self.tab_apv, "  Instituciones APV")
        self.tabs.addTab(self.tab_salud, "🏥  Instituciones Salud")

        # ── Pestaña Liquidaciones ──
        self.tab_liquidaciones = LiquidacionesTab()
        liq_lay = QVBoxLayout(self.panel_liq)
        liq_lay.setContentsMargins(0, 0, 0, 0); liq_lay.setSpacing(0)
        liq_lay.addWidget(self.tab_liquidaciones)

        self.tab_libro = LibroRemuneracionesTab()
        libro_lay = QVBoxLayout(self.panel_libro_main)
        libro_lay.setContentsMargins(0, 0, 0, 0); libro_lay.setSpacing(0)
        libro_lay.addWidget(self.tab_libro)

        self.panel_lre = QWidget()
        self.panel_lre.setStyleSheet("background:white;")
        self.tabs_main.insertTab(2, self.panel_lre, "📋  Libro Remun. Electrónico")
        self.tab_lre = LibroLRETab()
        lre_lay = QVBoxLayout(self.panel_lre)
        lre_lay.setContentsMargins(0, 0, 0, 0); lre_lay.setSpacing(0)
        lre_lay.addWidget(self.tab_lre)

        # Botón Salir
        btn_salir_layout = QHBoxLayout()
        btn_salir_layout.addStretch()
        self.btn_salir = QPushButton("🚪  Salir")
        self.btn_salir.setObjectName("btn_salir")
        self.btn_salir.setFixedHeight(38)
        self.btn_salir.setFixedWidth(120)
        self.btn_salir.clicked.connect(self._on_salir)
        btn_salir_layout.addWidget(self.btn_salir)
        layout.addLayout(btn_salir_layout)


        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("✅ Sistema iniciado correctamente")

    def _on_salir(self):
        resp = QMessageBox.question(
            self, "Confirmar salida",
            "¿Estás seguro de que deseas cerrar el programa?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if resp == QMessageBox.StandardButton.Yes:
            QApplication.quit()

    # ── Parámetros ───────────────────────────

    def _load_parametros(self):
        search = self.tab_param.search_box.text().strip()
        rows   = param_fetch_all(search)
        self.tab_param.table.setRowCount(len(rows))
        for r_idx, row in enumerate(rows):
            for c_idx, val in enumerate(row):
                if val is None:
                    text = ""
                elif c_idx in (0, 16):
                    text = str(val)
                elif c_idx == 13:
                    text = str(int(val))
                else:
                    try:
                        f = float(val)
                        text = f"{int(f):,}" if f == int(f) else f"{f:,.4f}".rstrip("0").rstrip(".")
                    except Exception:
                        text = str(val)
                item = QTableWidgetItem(text)
                item.setTextAlignment(
                    Qt.AlignmentFlag.AlignCenter if c_idx in (0, 13, 16)
                    else Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                )
                self.tab_param.table.setItem(r_idx, c_idx, item)
        self.status.showMessage(f"📅 Parámetros: {len(rows)} registro(s)")



    def _param_exportar_excel(self):
        """Exporta todos los parámetros mensuales a Excel."""
        import os
        from PyQt6.QtWidgets import QFileDialog
        ruta, _ = QFileDialog.getSaveFileName(
            self, "Guardar Excel",
            "C:/proyecto_Ia/parametros_mensuales_export.xlsx",
            "Excel (*.xlsx)"
        )
        if not ruta:
            return
        try:
            total = exportar_parametros_excel(ruta)
            self.status.showMessage(f"✅ Exportados {total} registros a Excel correctamente")
            QMessageBox.information(self, "Exportación exitosa",
                f"Se exportaron {total} registros correctamente.\n\nArchivo guardado en:\n{ruta}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo exportar:\n{e}")

    def _param_cargar_pdf(self):
        """Carga un PDF de Indicadores Previred y actualiza o crea el registro."""
        ruta, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar PDF de Indicadores Previred",
            "C:/proyecto_Ia",
            "Archivos PDF (*.pdf)"
        )
        if not ruta:
            return

        # Extraer mes/año del nombre del archivo
        mes_proc = extraer_mes_anio_pdf(ruta)
        if not mes_proc:
            QMessageBox.warning(self, "Nombre inválido",
                "No se pudo extraer el mes y año del nombre del archivo.\n"
                "Asegúrese de que el archivo se llame como:\n"
                "Indicadores-Previsionales-Previred-Abril-2026-1.pdf")
            return

        # Extraer datos del PDF
        datos, error = extraer_datos_pdf(ruta)
        if error:
            QMessageBox.critical(self, "Error al leer PDF", f"No se pudo leer el PDF:\n{error}")
            return

        # Verificar si el mes ya existe en la DB
        registro_existente = param_get(mes_proc)

        if registro_existente:
            resp = QMessageBox.question(
                self, "Mes ya existe",
                f"El mes '{mes_proc}' ya existe en la base de datos.\n\n"
                f"¿Desea ACTUALIZAR el registro con los datos del PDF?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if resp != QMessageBox.StandardButton.Yes:
                return
            accion = "actualizado"
        else:
            resp = QMessageBox.question(
                self, "Nuevo registro",
                f"El mes '{mes_proc}' no existe en la base de datos.\n\n"
                f"¿Desea CREAR un nuevo registro con los datos del PDF?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if resp != QMessageBox.StandardButton.Yes:
                return
            accion = "creado"

        # Calcular campos derivados
        uf          = datos['uf_mes']
        tope_imp_uf = datos['tope_imp_uf_afp']
        tope_ces_uf = datos['tope_ces_uf']
        sis         = datos['sis']
        ccaf        = datos['aporte_ccaf']
        calc        = calcular_campos(uf, tope_imp_uf, tope_ces_uf, sis, ccaf, mes_proc)

        # Determinar último día del mes y fecha formato
        ult_dia = ult_dia_para_mes(mes_proc)
        anio, mm = mes_proc.split('-')
        fecha_fmt = f"{ult_dia:02d}-{mm}-{anio}"

        values = [
            mes_proc,
            uf,
            tope_imp_uf,
            calc['tope_imp_pesos_afp'],
            tope_ces_uf,
            calc['tope_ces_pesos'],
            sis,
            calc['factor_sis'],
            calc['tope_salud_uf'],
            calc['tope_salud_pesos'],
            datos['imm'],
            datos['tope_gratif'],
            datos['monto_utm'],
            ult_dia,
            ccaf,
            calc['aporte_fonasa'],
            fecha_fmt,
            datos['aporte_afp'],
            datos['seg_social_exp_vida'],
        ]

        try:
            if accion == "actualizado":
                param_update(mes_proc, values[1:])
            else:
                param_insert(values)
            self._load_parametros()
            self.status.showMessage(f"✅ Registro '{mes_proc}' {accion} correctamente desde PDF")
            QMessageBox.information(self, "Éxito",
                f"El registro del mes '{mes_proc}' fue {accion} correctamente.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo guardar el registro:\n{e}")

    def _param_add(self):
        dlg = ParamDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            values = dlg.get_values()
            if param_get(values[0]):
                QMessageBox.warning(self, "Duplicado", f"Ya existe el mes '{values[0]}'."); return
            try:
                param_insert(values); self._load_parametros()
                self.status.showMessage(f"✅ Parámetro '{values[0]}' agregado")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _param_edit(self):
        mes = self.tab_param.selected_key(0)
        if not mes:
            QMessageBox.information(self, "Selección requerida", "Selecciona un registro para editar."); return
        record = param_get(mes)
        if not record: return
        dlg = ParamDialog(self, record=record)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            try:
                param_update(mes, dlg.get_values()[1:]); self._load_parametros()
                self.status.showMessage(f"✅ Parámetro '{mes}' actualizado")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _param_delete(self):
        mes = self.tab_param.selected_key(0)
        if not mes:
            QMessageBox.information(self, "Selección requerida", "Selecciona un registro para eliminar."); return
        if QMessageBox.question(self, "Confirmar", f"¿Eliminar el parámetro '{mes}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            try:
                param_delete(mes); self._load_parametros()
                self.status.showMessage(f"🗑️ Parámetro '{mes}' eliminado")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    # ── AFP ──────────────────────────────────

    def _load_afp(self):
        search   = self.tab_afp.search_box.text().strip()
        rows     = afp_fetch_all(search)
        previred = {cod for cod, _ in previred_get_all()}
        self.tab_afp.table.setRowCount(len(rows))
        for r_idx, row in enumerate(rows):
            for c_idx, val in enumerate(row):
                text = "" if val is None else str(val)
                item = QTableWidgetItem(text)
                # Marcar en rojo los códigos Previred no válidos
                if c_idx == 4 and text and text not in previred:
                    item.setForeground(QColor("#dc2626"))
                    item.setBackground(QColor("#fee2e2"))
                item.setTextAlignment(
                    Qt.AlignmentFlag.AlignCenter if c_idx in (0, 1, 3, 4)
                    else Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                    if c_idx == 5 else Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
                )
                self.tab_afp.table.setItem(r_idx, c_idx, item)
        self.status.showMessage(f"🏦 AFPs: {len(rows)} registro(s)")

    def _afp_add(self):
        dlg = AfpDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            values = dlg.get_values()
            if afp_get(values[0]):
                QMessageBox.warning(self, "Duplicado", f"Ya existe una AFP con el ID '{values[0]}'."); return
            existing_cod = afp_get_by_cod_prev(values[4])
            if existing_cod:
                QMessageBox.warning(self, "Codigo Previred duplicado",
                    f"El codigo Previred '{values[4]}' ya esta asignado a '{existing_cod[2]}'. "
                    f"No se pueden tener dos AFPs con el mismo codigo Previred."); return
            try:
                afp_insert(values); self._load_afp()
                self.status.showMessage(f"✅ AFP '{values[2]}' agregada")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _afp_edit(self):
        id_afp = self.tab_afp.selected_key(0)
        if not id_afp:
            QMessageBox.information(self, "Selección requerida", "Selecciona una AFP para editar."); return
        record = afp_get(id_afp)
        if not record: return
        dlg = AfpDialog(self, record=record)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            values = dlg.get_values()
            existing_cod = afp_get_by_cod_prev(values[4], exclude_id=id_afp)
            if existing_cod:
                QMessageBox.warning(self, "Codigo Previred duplicado",
                    f"El codigo Previred '{values[4]}' ya esta asignado a '{existing_cod[2]}'. "
                    f"No se pueden tener dos AFPs con el mismo codigo Previred."); return
            try:
                afp_update(id_afp, values[1:]); self._load_afp()
                self.status.showMessage(f"✅ AFP '{id_afp}' actualizada")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _afp_delete(self):
        id_afp = self.tab_afp.selected_key(0)
        if not id_afp:
            QMessageBox.information(self, "Selección requerida", "Selecciona una AFP para eliminar."); return
        if QMessageBox.question(self, "Confirmar", f"¿Eliminar la AFP '{id_afp}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            try:
                afp_delete(id_afp); self._load_afp()
                self.status.showMessage(f"🗑️ AFP '{id_afp}' eliminada")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

    def _load_salud(self):
        search = self.tab_salud.search_box.text().strip()
        rows = salud_fetch_all(search)
        self.tab_salud.table.setRowCount(len(rows))
        for r_idx, row in enumerate(rows):
            for c_idx, val in enumerate(row):
                text = "" if val is None else str(val)
                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tab_salud.table.setItem(r_idx, c_idx, item)
        self.status.showMessage(f"Salud: {len(rows)} registro(s)")

    def _load_empresas(self):
        search=self.tab_empresas.search_box.text().strip()
        rows=empresas_fetch_all(search)
        self.tab_empresas.table.setRowCount(len(rows))
        for r_idx,row in enumerate(rows):
            for c_idx,val in enumerate(row):
                item=QTableWidgetItem('' if val is None else str(val))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tab_empresas.table.setItem(r_idx,c_idx,item)
        self.status.showMessage(f'Empresas: {len(rows)} registro(s)')

    def _empresas_add(self):
        dlg=EmpresaDialog(self)
        if dlg.exec()==QDialog.DialogCode.Accepted:
            values=dlg.get_values()
            if empresas_get(values[0]):
                QMessageBox.warning(self,'Duplicado',f'Ya existe {values[0]}');return
            try:
                empresas_insert(values);self._load_empresas()
            except Exception as e:
                QMessageBox.critical(self,'Error',str(e))

    def _empresas_edit(self):
        id_emp=self.tab_empresas.selected_key(0)
        if not id_emp:return
        record=empresas_get(id_emp)
        if not record:return
        dlg=EmpresaDialog(self,record=record)
        if dlg.exec()==QDialog.DialogCode.Accepted:
            try:
                empresas_update(id_emp,dlg.get_values()[1:]);self._load_empresas()
            except Exception as e:
                QMessageBox.critical(self,'Error',str(e))

    def _empresas_delete(self):
        id_emp=self.tab_empresas.selected_key(0)
        if not id_emp:return
        if QMessageBox.question(self,'Confirmar',f'Eliminar empresa {id_emp}?',QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)==QMessageBox.StandardButton.Yes:
            try:
                empresas_delete(id_emp);self._load_empresas()
            except Exception as e:
                QMessageBox.critical(self,'Error',str(e))

    def _load_cajas(self):
        search = self.tab_cajas.search_box.text().strip()
        rows = cajas_fetch_all(search)
        self.tab_cajas.table.setRowCount(len(rows))
        for r_idx, row in enumerate(rows):
            for c_idx, val in enumerate(row):
                item = QTableWidgetItem("" if val is None else str(val))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tab_cajas.table.setItem(r_idx, c_idx, item)
        self.status.showMessage(f"Cajas: {len(rows)} registro(s)")

    def _cajas_add(self):
        dlg = CajasDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            values = dlg.get_values()
            if cajas_get(values[0]):
                QMessageBox.warning(self, "Duplicado", f"Ya existe {values[0]}"); return
            try:
                cajas_insert(values); self._load_cajas()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _cajas_edit(self):
        id_inst = self.tab_cajas.selected_key(0)
        if not id_inst: return
        record = cajas_get(id_inst)
        if not record: return
        dlg = CajasDialog(self, record=record)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            try:
                cajas_update(id_inst, dlg.get_values()[1:]); self._load_cajas()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _cajas_delete(self):
        id_inst = self.tab_cajas.selected_key(0)
        if not id_inst: return
        if QMessageBox.question(self, "Confirmar", f"Eliminar {id_inst}?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            try:
                cajas_delete(id_inst); self._load_cajas()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _load_mutuales(self):
        search = self.tab_mutuales.search_box.text().strip()
        rows = mutuales_fetch_all(search)
        self.tab_mutuales.table.setRowCount(len(rows))
        for r_idx, row in enumerate(rows):
            for c_idx, val in enumerate(row):
                item = QTableWidgetItem("" if val is None else str(val))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tab_mutuales.table.setItem(r_idx, c_idx, item)
        self.status.showMessage(f"Mutuales: {len(rows)} registro(s)")

    def _mutuales_add(self):
        dlg = MutualesDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            values = dlg.get_values()
            if mutuales_get(values[0]):
                QMessageBox.warning(self, "Duplicado", f"Ya existe {values[0]}"); return
            try:
                mutuales_insert(values); self._load_mutuales()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _mutuales_edit(self):
        id_inst = self.tab_mutuales.selected_key(0)
        if not id_inst: return
        record = mutuales_get(id_inst)
        if not record: return
        dlg = MutualesDialog(self, record=record)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            try:
                mutuales_update(id_inst, dlg.get_values()[1:]); self._load_mutuales()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _mutuales_delete(self):
        id_inst = self.tab_mutuales.selected_key(0)
        if not id_inst: return
        if QMessageBox.question(self, "Confirmar", f"Eliminar {id_inst}?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            try:
                mutuales_delete(id_inst); self._load_mutuales()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _load_apv(self):
        search = self.tab_apv.search_box.text().strip()
        rows = apv_fetch_all(search)
        self.tab_apv.table.setRowCount(len(rows))
        for r_idx, row in enumerate(rows):
            for c_idx, val in enumerate(row):
                text = "" if val is None else str(val)
                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tab_apv.table.setItem(r_idx, c_idx, item)
        self.status.showMessage(f"APV: {len(rows)} registro(s)")

    def _apv_add(self):
        dlg = ApvDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            values = dlg.get_values()
            if apv_get(values[0]):
                QMessageBox.warning(self, "Duplicado", f"Ya existe {values[0]}"); return
            try:
                apv_insert(values); self._load_apv()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _apv_edit(self):
        id_apv = self.tab_apv.selected_key(0)
        if not id_apv: return
        record = apv_get(id_apv)
        if not record: return
        dlg = ApvDialog(self, record=record)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            try:
                apv_update(id_apv, dlg.get_values()[1:]); self._load_apv()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _apv_delete(self):
        id_apv = self.tab_apv.selected_key(0)
        if not id_apv: return
        if QMessageBox.question(self, "Confirmar", f"Eliminar {id_apv}?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            try:
                apv_delete(id_apv); self._load_apv()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _salud_add(self):
        dlg = SaludDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            values = dlg.get_values()
            if salud_get(values[0]):
                QMessageBox.warning(self, "Duplicado", f"Ya existe {values[0]}"); return
            try:
                salud_insert(values); self._load_salud()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _salud_edit(self):
        id_inst = self.tab_salud.selected_key(0)
        if not id_inst: return
        record = salud_get(id_inst)
        if not record: return
        dlg = SaludDialog(self, record=record)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            try:
                salud_update(id_inst, dlg.get_values()[1:]); self._load_salud()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _salud_delete(self):
        id_inst = self.tab_salud.selected_key(0)
        if not id_inst: return
        if QMessageBox.question(self, "Confirmar", f"Eliminar {id_inst}?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            try:
                salud_delete(id_inst); self._load_salud()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

if __name__ == "__main__":
    init_db()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(STYLE)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
