"""
Rex+ Tools — Minutas de Implementación
Formulario para generar minutas de Remuneración y Asistencia en Excel.
"""

import io
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    from lib.branding import aplicar_branding, aplicar_footer, hero
    BRANDING = True
except ImportError:
    BRANDING = False

st.set_page_config(page_title="Minutas | Rex+ Tools", page_icon="📋", layout="wide")

if BRANDING:
    aplicar_branding(titulo_pagina="Minutas", badge="PRODUCCIÓN")
    hero("📋 Minutas de Implementación", "Completa los datos y descarga la minuta en Excel lista para entregar.")
else:
    st.title("📋 Minutas de Implementación")
    st.caption("Completa los datos y descarga la minuta en Excel lista para entregar.")

COLOR_TITLE    = "1B3A6B"
COLOR_SECTION  = "2E6DB4"
COLOR_ROW_ODD  = "D9E8F5"
COLOR_ROW_EVEN = "EEF5FB"
COLOR_CHECK1   = "F0F7E6"
COLOR_CHECK2   = "E8F5E9"
WHITE          = "FFFFFF"
FONT_WHITE     = "FFFFFF"
FONT_DARK      = "1A3A5F"

SI_NO  = ["si", "no", "No aplica"]
SI_NO2 = ["si", "no"]

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color=FONT_DARK, size=11):
    return Font(name="Arial", bold=bold, color=color, size=size)

def _border():
    thin = Side(style="thin", color="BBCDE0")
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def _set_title(ws, row, text):
    cell = ws.cell(row=row, column=3, value=text)
    cell.fill = _fill(COLOR_TITLE)
    cell.font = Font(name="Arial", bold=True, color=FONT_WHITE, size=13)
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=5)
    ws.row_dimensions[row].height = 28

def _set_section(ws, row, text):
    cell = ws.cell(row=row, column=3, value=text)
    cell.fill = _fill(COLOR_SECTION)
    cell.font = Font(name="Arial", bold=True, color=FONT_WHITE, size=11)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=5)
    ws.row_dimensions[row].height = 22

def _set_row(ws, row, label, value, even=False):
    bg_label = COLOR_ROW_EVEN if even else COLOR_ROW_ODD
    bg_value = COLOR_ROW_EVEN if even else WHITE
    lc = ws.cell(row=row, column=3, value=label)
    lc.fill = _fill(bg_label); lc.font = _font()
    lc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    lc.border = _border()
    vc = ws.cell(row=row, column=4, value=value)
    vc.fill = _fill(bg_value); vc.font = _font(color="1A3A5F")
    vc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    vc.border = _border()
    ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=5)
    ws.row_dimensions[row].height = 20

def _set_check(ws, row, text, alt=False):
    bg = COLOR_CHECK2 if alt else COLOR_CHECK1
    c1 = ws.cell(row=row, column=3, value="☐ Pendiente")
    c1.fill = _fill(bg); c1.font = Font(name="Arial", bold=True, color="4CAF50", size=11)
    c1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c1.border = _border()
    c2 = ws.cell(row=row, column=4, value=text)
    c2.fill = _fill(bg); c2.font = _font()
    c2.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c2.border = _border()
    ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=5)
    ws.row_dimensions[row].height = 20

def _set_note(ws, row):
    c = ws.cell(row=row, column=3,
                value="💡  Los campos en azul son editables. Use los desplegables para seleccionar valores estándar.")
    c.font = Font(name="Arial", italic=True, color="555555", size=9)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=5)
    ws.row_dimensions[row].height = 18

def _col_widths(ws):
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 3
    ws.column_dimensions["C"].width = 38
    ws.column_dimensions["D"].width = 35
    ws.column_dimensions["E"].width = 5

def build_remuneraciones(ws, d):
    _col_widths(ws)
    _set_title(ws, 2, "📋  MINUTA DE IMPLEMENTACIÓN — REX REMUNERACIONES")
    _set_section(ws, 4, "  DATOS GENERALES DEL CLIENTE")
    fields = [
        ("OT (Orden de Trabajo)", d["ot"]),
        ("Empresa / Razón Social", d["empresa"]),
        ("RUT Empresa", d["rut"]),
        ("Vendedor", d["vendedor"]),
        ("Jefe de Proyecto / Dirección", d["direccion"]),
        ("Correo de Contacto", d["correo"]),
        ("Número de Contacto", d["telefono"]),
        ("Plan Contratado", d["plan"]),
        ("Cantidad de Colaboradores", d["colaboradores"]),
        ("Cantidad de Razones Sociales", d["razones_sociales"]),
    ]
    for i, (l, v) in enumerate(fields):
        _set_row(ws, 5+i, l, v, even=(i%2==1))
    _set_section(ws, 16, "  CONFIGURACIÓN DE REMUNERACIONES")
    cfg = [
        ("Estructura de Remuneraciones", d["estructura"]),
        ("Comisión / Semana Corrida", d["comision"]),
        ("Reliquidación / Renta Accesoria", d["reliquidacion"]),
        ("3 Primeros Días (Art. 195)", d["tres_dias"]),
        ("Zona Extrema", d["zona_extrema"]),
        ("Provisión Vacaciones", d["provision"]),
        ("Centralización Contable", d["centralizacion"]),
        ("Transferencia Bancaria", d["transferencia"]),
    ]
    for i, (l, v) in enumerate(cfg):
        _set_row(ws, 17+i, l, v, even=(i%2==1))
    _set_section(ws, 26, "  COMENTARIOS GENERALES")
    _set_row(ws, 27, "Minuta / Observaciones", d["observaciones"])
    ws.row_dimensions[27].height = 40
    _set_section(ws, 30, "  ✅  CHECKLIST — INFORMACIÓN NECESARIA PARA COMENZAR")
    checklist = [
        "Liquidaciones de todo 2026",
        "Libro de remuneraciones (en Excel) o el que sube a la DT",
        "Contratos y finiquitos (en Word)",
        "Saldo y vacaciones (en Excel)",
        "Licencias médicas (Registros)",
        "Ausentismos (en Excel)",
    ]
    for i, item in enumerate(checklist):
        _set_check(ws, 31+i, item, alt=(i%2==1))
    _set_note(ws, 38)

def build_asistencia(ws, d):
    _col_widths(ws)
    _set_title(ws, 2, "📋  MINUTA DE IMPLEMENTACIÓN — REX ASISTENCIA")
    _set_section(ws, 4, "  DATOS DE LA EMPRESA")
    emp = [
        ("Empresa", d["empresa"]),
        ("RUT", d["rut"]),
        ("Dirección", d["direccion"]),
        ("Vendedor", d["vendedor"]),
        ("Empresa Venta", d["empresa_venta"]),
        ("Contacto (Nombre Completo)", d["contacto_nombre"]),
        ("Contacto (Número)", d["contacto_numero"]),
        ("Contacto (Email)", d["contacto_email"]),
    ]
    for i, (l, v) in enumerate(emp):
        _set_row(ws, 5+i, l, v, even=(i%2==1))
    _set_section(ws, 14, "  PLAN DE IMPLEMENTACIÓN")
    _set_row(ws, 15, "Plan Asistencia", d["plan_asistencia"])
    _set_row(ws, 16, "Plan Casino", d["plan_casino"], even=True)
    _set_row(ws, 17, "Adicionales", d["adicionales_plan"])
    _set_section(ws, 20, "  CONSULTAS TÉCNICAS")
    tec = [
        ("Sistema de Asistencia Actual", d["sistema_actual"]),
        ("Dispositivo de Marcaje", d["dispositivo"]),
        ("¿Tiene Rex+?", d["tiene_rex"]),
        ("Cantidad de RUT", d["cantidad_rut"]),
        ("Cantidad de Empleados", d["cantidad_empleados"]),
        ("Empleados Art. 22", d["art22"]),
        ("Tipos de Horario", d["tipos_horario"]),
        ("Cantidad de Ubicaciones", d["ubicaciones"]),
    ]
    for i, (l, v) in enumerate(tec):
        _set_row(ws, 21+i, l, v, even=(i%2==1))
    _set_section(ws, 30, "  CONFIGURACIÓN ADICIONAL")
    _set_row(ws, 31, "Concepto de Asistencia (Remuneración)", d["concepto"])
    _set_row(ws, 32, "Cortes Mensuales", d["cortes"], even=True)
    _set_row(ws, 33, "Adicionales / Observaciones", d["observaciones"])
    _set_note(ws, 36)

def generar_excel(data_rem, data_asi):
    wb = Workbook()
    ws_rem = wb.active
    ws_rem.title = "Rex - Remuneraciones"
    build_remuneraciones(ws_rem, data_rem)
    ws_asi = wb.create_sheet("Rex - Asistencia")
    build_asistencia(ws_asi, data_asi)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────

tab_rem, tab_asi = st.tabs(["📊 Remuneraciones", "🕐 Asistencia"])

with tab_rem:
    st.subheader("Datos Generales del Cliente")
    c1, c2 = st.columns(2)
    ot              = c1.text_input("OT (Orden de Trabajo)", placeholder="Ej: 2757", key="r_ot")
    empresa_r       = c1.text_input("Empresa / Razón Social", placeholder="Ej: Fundación Ejemplo", key="r_empresa")
    rut_r           = c1.text_input("RUT Empresa", placeholder="Ej: 65058734-0", key="r_rut")
    vendedor_r      = c1.text_input("Vendedor", placeholder="Ej: Matias Ossandon", key="r_vendedor")
    direccion_r     = c2.text_input("Jefe de Proyecto / Dirección", placeholder="Ej: Av. Principal 123", key="r_direccion")
    correo_r        = c2.text_input("Correo de Contacto", placeholder="correo@empresa.cl", key="r_correo")
    telefono_r      = c2.text_input("Número de Contacto", placeholder="Ej: 56912345678", key="r_telefono")
    plan_r          = c2.selectbox("Plan Contratado", ["estandar", "premium", "basico", "otro"], key="r_plan")
    col_r, col_rs   = st.columns(2)
    colaboradores_r = col_r.number_input("Cantidad de Colaboradores", min_value=1, value=1, step=1, key="r_colab")
    razones_r       = col_rs.number_input("Cantidad de Razones Sociales", min_value=1, value=1, step=1, key="r_razones")

    st.divider()
    st.subheader("Configuración de Remuneraciones")
    c3, c4 = st.columns(2)
    estructura_r     = c3.selectbox("Estructura de Remuneraciones",
                                    ["Fijos y variables", "Solo fijos", "Solo variables", "Otro"], key="r_estructura")
    comision_r       = c3.selectbox("Comisión / Semana Corrida", SI_NO, key="r_comision")
    reliquidacion_r  = c3.selectbox("Reliquidación / Renta Accesoria", SI_NO, key="r_reliq")
    tres_dias_r      = c3.selectbox("3 Primeros Días (Art. 195)", ["No aplica", "si", "no"], key="r_tresdias")
    zona_extrema_r   = c4.selectbox("Zona Extrema", SI_NO2, key="r_zona")
    provision_r      = c4.selectbox("Provisión Vacaciones", SI_NO2, key="r_provision")
    centralizacion_r = c4.selectbox("Centralización Contable",
                                    ["defontana", "sap", "oracle", "otro sistema", "no aplica"], key="r_central")
    transferencia_r  = c4.text_input("Transferencia Bancaria",
                                     placeholder="Ej: banco chile y banco estado", key="r_transfer")

    st.divider()
    st.subheader("Comentarios Generales")
    observaciones_r = st.text_area("Minuta / Observaciones", height=80,
                                   placeholder="Ej: SB+ no tiene gratificación + Colación y movilización",
                                   key="r_obs")

with tab_asi:
    st.subheader("Datos de la Empresa")
    c5, c6 = st.columns(2)
    empresa_a       = c5.text_input("Empresa", placeholder="Ej: Municipalidad de Marchigue", key="a_empresa")
    rut_a           = c5.text_input("RUT", placeholder="Ej: 69091300-3", key="a_rut")
    direccion_a     = c5.text_input("Dirección", placeholder="Ej: Maria Errazuriz 1507", key="a_direccion")
    vendedor_a      = c5.text_input("Vendedor", placeholder="Ej: Matias Ossandon", key="a_vendedor")
    empresa_venta_a = c6.selectbox("Empresa Venta", ["REX", "Visma", "Otro"], key="a_emp_venta")
    contacto_nombre = c6.text_input("Contacto (Nombre Completo)", placeholder="Nombre del contacto", key="a_cont_nombre")
    contacto_numero = c6.text_input("Contacto (Número)", placeholder="Ej: 56912345678", key="a_cont_num")
    contacto_email  = c6.text_input("Contacto (Email)", placeholder="correo@empresa.cl", key="a_cont_email")

    st.divider()
    st.subheader("Plan de Implementación")
    plan_a = st.text_input("Plan Asistencia",
                           placeholder="Ej: PLAN ASISTENCIA CON MARCAJE Reloj, APP Y/O WEB", key="a_plan")
    c7, c8 = st.columns(2)
    casino_a           = c7.text_input("Plan Casino", value="NO APLICA", key="a_casino")
    adicionales_plan_a = c8.text_input("Adicionales", value="NO APLICA", key="a_adicionales_plan")

    st.divider()
    st.subheader("Consultas Técnicas")
    c9, c10 = st.columns(2)
    sistema_a     = c9.text_input("Sistema de Asistencia Actual", placeholder="Ej: Cass, Manual", key="a_sistema")
    dispositivo_a = c9.text_input("Dispositivo de Marcaje", placeholder="Ej: APP y Reloj control", key="a_dispositivo")
    tiene_rex_a   = c9.selectbox("¿Tiene Rex+?", ["No", "Si"], key="a_tiene_rex")
    cant_rut_a    = c9.number_input("Cantidad de RUT", min_value=1, value=1, step=1, key="a_cant_rut")
    cant_emp_a    = c10.number_input("Cantidad de Empleados", min_value=1, value=1, step=1, key="a_cant_emp")
    art22_a       = c10.selectbox("Empleados Art. 22", ["No", "Si", "Parcial"], key="a_art22")
    horario_a     = c10.text_input("Tipos de Horario", placeholder="Ej: Varios, Turno fijo", key="a_horario")
    ubicaciones_a = c10.text_input("Cantidad de Ubicaciones", placeholder="Ej: 1, Varias", key="a_ubicaciones")

    st.divider()
    st.subheader("Configuración Adicional")
    c11, c12 = st.columns(2)
    concepto_a      = c11.text_input("Concepto de Asistencia (Remuneración)",
                                     placeholder="Ej: Horas atraso y extra", key="a_concepto")
    cortes_a        = c11.text_input("Cortes Mensuales", placeholder="Ej: 24c/m", key="a_cortes")
    observaciones_a = c12.text_area("Adicionales / Observaciones", height=80,
                                    placeholder="Observaciones adicionales...", key="a_obs")

# ── Descarga ─────────────────────────────────────────────────────────────────
st.divider()
st.subheader("Descargar Minutas")

nombre_archivo = st.text_input(
    "Nombre del archivo (sin extensión)",
    value=f"Minuta_{empresa_r or empresa_a or 'cliente'}".replace(" ", "_"),
    key="nombre_archivo",
)

if st.button("📥 Generar y Descargar Excel", type="primary", use_container_width=False):
    data_rem = {
        "ot": ot, "empresa": empresa_r, "rut": rut_r,
        "vendedor": vendedor_r, "direccion": direccion_r,
        "correo": correo_r, "telefono": telefono_r,
        "plan": plan_r, "colaboradores": colaboradores_r,
        "razones_sociales": razones_r, "estructura": estructura_r,
        "comision": comision_r, "reliquidacion": reliquidacion_r,
        "tres_dias": tres_dias_r, "zona_extrema": zona_extrema_r,
        "provision": provision_r, "centralizacion": centralizacion_r,
        "transferencia": transferencia_r, "observaciones": observaciones_r,
    }
    data_asi = {
        "empresa": empresa_a, "rut": rut_a, "direccion": direccion_a,
        "vendedor": vendedor_a, "empresa_venta": empresa_venta_a,
        "contacto_nombre": contacto_nombre, "contacto_numero": contacto_numero,
        "contacto_email": contacto_email, "plan_asistencia": plan_a,
        "plan_casino": casino_a, "adicionales_plan": adicionales_plan_a,
        "sistema_actual": sistema_a, "dispositivo": dispositivo_a,
        "tiene_rex": tiene_rex_a, "cantidad_rut": cant_rut_a,
        "cantidad_empleados": cant_emp_a, "art22": art22_a,
        "tipos_horario": horario_a, "ubicaciones": ubicaciones_a,
        "concepto": concepto_a, "cortes": cortes_a,
        "observaciones": observaciones_a,
    }
    excel_bytes = generar_excel(data_rem, data_asi)
    st.download_button(
        label="⬇️ Haz clic aquí para descargar",
        data=excel_bytes,
        file_name=f"{nombre_archivo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="descarga_excel",
    )
    st.success("✅ Minuta generada. Haz clic en el botón azul para descargar.")

if BRANDING:
    aplicar_footer()
