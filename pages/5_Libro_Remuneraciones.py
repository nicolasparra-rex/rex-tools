"""
Rex+ Tools - Libro de Remuneraciones Electrónico
Transforma el CSV del LRE al formato de importación Rex+.
"""

import streamlit as st
import pandas as pd
import math
import os
import io
import sys
from pathlib import Path

# ── Paths ─────────────────────────────────────────────────────────────────────
_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(_ROOT))
sys.path.insert(0, str(_ROOT / "lib"))

from lib.branding import aplicar_branding, aplicar_footer, hero

st.set_page_config(
    page_title="Migración Historia LRE | Rex+ Tools",
    page_icon="📊",
    layout="wide",
)

aplicar_branding(titulo_pagina="Libro de Remuneraciones")
hero(
    titulo="Migración Historia mediante LRE",
    descripcion="Transforma el CSV del LRE al formato de importación Rex+. Carga los archivos de referencia y el CSV para comenzar.",
    icono="📊",
)

# ── Ruta equivalencias base (repo) ────────────────────────────────────────────
EQUIVALENCIAS_BASE = _ROOT / "Equivalencia" / "equivalencias.xlsx"
# Fallback si no está en Equivalencia/ (archivos en raíz del repo)
if not EQUIVALENCIAS_BASE.exists():
    EQUIVALENCIAS_BASE = _ROOT / "equivalencias.xlsx"

# ── Session state ─────────────────────────────────────────────────────────────
for k, v in {
    "empleado_bytes":    None,
    "empresa_bytes":     None,
    "equiv_override":    None,
    "refs_cargadas":     False,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ── Helpers ───────────────────────────────────────────────────────────────────
def normalizar_rut(rut: str) -> str:
    return str(rut).strip().upper()

def rut_empresa_desde_filename(nombre: str) -> str:
    partes = os.path.splitext(nombre)[0].split("_")
    raw = partes[-2]
    return raw[:-1] + "-" + raw[-1]

def periodo_desde_filename(nombre: str) -> str:
    partes = os.path.splitext(nombre)[0].split("_")
    raw = partes[-1]
    return f"{raw[:4]}-{raw[4:]}"

def safe_get(d, key, default=None):
    return d.get(key, d.get(str(key), default))

def round_awz(x):
    return math.floor(x + 0.5) if x >= 0 else math.ceil(x - 0.5)


# ── Carga de referencias ──────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def cargar_referencias(equiv_bytes, empleado_bytes, empresa_bytes):
    xl_eq = pd.ExcelFile(io.BytesIO(equiv_bytes))
    df_afp     = pd.read_excel(xl_eq, "afp")
    df_porafp  = pd.read_excel(xl_eq, "porafp")
    df_salud   = pd.read_excel(xl_eq, "salud")
    df_sis     = pd.read_excel(xl_eq, "sis")
    df_ccaf    = pd.read_excel(xl_eq, "ccaf")
    df_porccaf = pd.read_excel(xl_eq, "porccaf")
    df_tijor   = pd.read_excel(xl_eq, "tijor")
    df_empleado = pd.read_excel(io.BytesIO(empleado_bytes), "Hoja 1")
    df_empresa  = pd.read_excel(io.BytesIO(empresa_bytes),  "Hoja 1")
    return {
        "afp": df_afp, "porafp": df_porafp, "salud": df_salud,
        "sis": df_sis, "ccaf": df_ccaf, "porccaf": df_porccaf,
        "tijor": df_tijor, "empleado": df_empleado, "empresa": df_empresa,
    }


# ── Transformación ────────────────────────────────────────────────────────────
def transformar(input_bytes, nombre, refs):
    rut_empresa = normalizar_rut(rut_empresa_desde_filename(nombre))
    periodo     = periodo_desde_filename(nombre)

    df_emp_tbl = refs["empresa"].copy()
    df_emp_tbl["Rut empresa"] = df_emp_tbl["Rut empresa"].str.upper()
    fila_emp = df_emp_tbl[df_emp_tbl["Rut empresa"] == rut_empresa]
    if fila_emp.empty:
        raise ValueError(f"RUT empresa '{rut_empresa}' no encontrado en empresa.xlsx")
    fila_emp   = fila_emp.iloc[0]
    id_empresa = fila_emp["idempresa"]
    mutual     = fila_emp["Mutual"]
    pct_mutual = fila_emp["% mutual"]

    sis_row  = refs["sis"][refs["sis"]["proceso"] == periodo]
    pct_sis  = float(sis_row["porcentaje"].iloc[0]) if not sis_row.empty else 0.0
    porccaf_row = refs["porccaf"][refs["porccaf"]["periodo"] == periodo]
    pct_ccaf = float(porccaf_row["porcentaje"].iloc[0]) if not porccaf_row.empty else 0.0
    porafp_periodo = (
        refs["porafp"][refs["porafp"]["proceso"] == periodo]
        .set_index("nombre")["porcentaje"].to_dict()
    )

    afp_map   = dict(zip(refs["afp"]["afp"], refs["afp"]["idrex"]))
    salud_map = {int(r["id"]): r["idrex"] for _, r in refs["salud"].iterrows() if pd.notna(r["id"])}
    ccaf_map  = {int(r["id"]): r["idrex"] for _, r in refs["ccaf"].iterrows() if pd.notna(r["id"])}
    tijor_map = dict(zip(refs["tijor"]["id"], refs["tijor"]["idrex"]))

    df_emp = refs["empleado"].copy()
    df_emp["Id empleado"] = df_emp["Id empleado"].str.upper()
    df_emp["_key"] = df_emp["Id empleado"] + "_" + df_emp["fechaInic"].dt.strftime("%Y-%m-%d").fillna("")
    df_emp_idx = df_emp.set_index("_key")

    df_in = pd.read_csv(io.BytesIO(input_bytes), encoding="latin-1", sep=";")
    df_in["Rut trabajador(1101)"] = df_in["Rut trabajador(1101)"].str.upper()

    filas = []
    for _, row in df_in.iterrows():
        rut_trab   = normalizar_rut(row["Rut trabajador(1101)"])
        afp_code   = int(row.get("AFP(1141)", 100))
        afp_name   = safe_get(afp_map, afp_code, "afp")
        pct_afp    = float(porafp_periodo.get(afp_name, 0) or 0)
        salud_code = int(row.get("FONASA - ISAPRE(1143)", 102))
        isapre     = safe_get(salud_map, salud_code, "fonasa")
        ccaf_code  = row.get("CCAF(1110)", 0)
        try:    ccaf_name = safe_get(ccaf_map, int(ccaf_code), "sincaja")
        except: ccaf_name = "sincaja"
        jorn_code = row.get("Código tipo de jornada(1107)", 101)
        try:    jornada = safe_get(tijor_map, int(jorn_code), "C")
        except: jornada = "C"

        fecha_inic_raw = row.get("Fecha inicio contrato(1102)", None)
        try:
            fecha_inic_lre = pd.to_datetime(fecha_inic_raw, dayfirst=True)
            fecha_inic_str = fecha_inic_lre.strftime("%Y-%m-%d")
        except:
            fecha_inic_lre = None
            fecha_inic_str = ""

        emp_key  = f"{rut_trab}_{fecha_inic_str}"
        emp_info = df_emp_idx.loc[emp_key] if emp_key in df_emp_idx.index else None
        if emp_info is not None:
            num_contrato = int(emp_info.get("Numero de contrato", 1) or 1)
            tipo_cont    = str(emp_info.get("tipoCont", None) or "").strip()
        else:
            num_contrato = "revisar"
            tipo_cont    = ""

        if fecha_inic_lre is not None:
            fecha_fin_mes = pd.Timestamp(f"{periodo}-01") + pd.offsets.MonthEnd(0)
            antiguedad = abs((fecha_fin_mes.year - fecha_inic_lre.year) * 12 + (fecha_fin_mes.month - fecha_inic_lre.month))
        else:
            antiguedad = 0

        afc_input     = float(row.get("AFC - Aporte empleador(4151)", 0) or 0)
        causal_raw    = row.get("Causal término de contrato(1104)", None)
        try:    causal = int(causal_raw)
        except: causal = None

        if antiguedad > 132:
            afc_solid_val = round_awz(afc_input)
        elif tipo_cont in ("O", "F"):
            afc_solid_val = round_awz(afc_input / 3 * 0.2)
        elif tipo_cont == "I":
            afc_solid_val = round_awz(afc_input / 2.4 * 0.8)
        elif (tipo_cont == "I" and causal == 6) or causal == 7:
            afc_solid_val = round_awz(afc_input / 3 * 0.2)
        else:
            afc_solid_val = 0

        if (tipo_cont == "I" and causal == 6) or causal == 7:
            afc_indi_val = round_awz(afc_input / 3 * 2.8)
        elif tipo_cont == "I" and antiguedad <= 132:
            afc_indi_val = round_awz(afc_input / 2.4 * 1.6)
        elif tipo_cont == "I" and antiguedad >= 133:
            afc_indi_val = 0
        elif tipo_cont in ("O", "F"):
            afc_indi_val = round_awz(afc_input / 3 * 2.8)
        else:
            afc_indi_val = 0

        if tipo_cont == "":
            dif_afc = "sin Tipo de contrato"
        elif (afc_solid_val + afc_indi_val) != afc_input:
            dif_afc = "revisar con cliente"
        else:
            dif_afc = ""

        def v(col, default=0):
            return row.get(col, default) or default

        val_liq = (
            v("Sueldo(2101)") + v("Sobresueldo(2102)") + v("Comisiones(2103)") +
            v("Semana corrida(2104)") + v("Participación(2105)") + v("Gratificación(2106)") +
            v("Recargo 30% día domingo(2107)") + v("Remun. variable pagada en vacaciones(2108)") +
            v("Aguinaldo(2110)") + v("Bonos u otras remun. fijas mensuales(2111)") +
            v("Tratos(2112)") + v("Bonos u otras remun. variables mensuales o superiores a un mes(2113)") +
            v("Beneficios en especie constitutivos de remun(2115)") +
            v("Otras remuneraciones superiores a un mes(2123)") +
            v("Pago por horas de trabajo sindical(2124)") +
            v("Subsidio por incapacidad laboral por licencia médica(2201)") +
            v("Beca de estudio(2202)") + v("Otros ingresos no constitutivos de renta(2204)") +
            v("Colación(2301)") + v("Movilización(2302)") + v("Viáticos(2303)") +
            v("Asignación de pérdida de caja(2304)") + v("Asignación de desgaste herramienta(2305)") +
            v("Gastos por causa del trabajo(2306)") + v("Sala cuna(2308)") +
            v("Alojamiento por razones de trabajo(2310)") + v("Asignación familiar legal(2311)") +
            v("Asignación trabajo a distancia o teletrabajo(2309)") + v("Asignación de traslación(2312)") +
            v("Indemnización por feriado legal(2313)") + v("Indemnización años de servicio(2314)") +
            v("Indemnización sustitutiva del aviso previo(2315)") + v("Indemnización fuero maternal(2316)") +
            v("Pago indemnización a todo evento(2331)") +
            v("Indemnizaciones voluntarias tributables(2417)") +
            v("Indemnizaciones contractuales tributables(2418)")
        ) - (
            v("Cotización obligatoria previsional (AFP o IPS)(3141)") +
            v("Cotización obligatoria salud 7%(3143)") + v("Cotización voluntaria para salud(3144)") +
            v("Cotización AFC - trabajador(3151)") +
            v("Cotización adicional trabajo pesado - trabajador(3154)") +
            v("Cotización APVi Mod A(3155)") + v("Cotización APVi Mod B hasta UF50(3156)") +
            v("Impuesto retenido por remuneraciones(3161)") +
            v("Impuesto retenido por indemnizaciones(3162)") +
            v("Mayor retención de impuestos solicitada por el trabajador(3163)") +
            v("Retención préstamo clase media 2020 (Ley 21.252) (3166)") +
            v("Cuota sindical 1(3171)") + v("Crédito social CCAF(3110)") +
            v("Cuota vivienda o educación(3181)") + v("Crédito cooperativas de ahorro(3182)") +
            v("Otros descuentos autorizados y solicitados por el trabajador(3183)") +
            v("Otros descuentos(3185)") + v("Pensiones de alimentos(3186)") +
            v("Descuentos por anticipos y préstamos(3188)")
        ) - v("Total líquido(5501)")

        filas.append({
            "Fecha de proceso": periodo, "Id empleado": rut_trab,
            "Número de contrato": num_contrato, "Id de empresa": id_empresa,
            "afp": afp_name, "%afp": pct_afp, "isapre": isapre,
            "Ccaf": ccaf_name, "% Ccaf": pct_ccaf, "Mutual": mutual,
            "% mutual": pct_mutual, "%Sis": pct_sis,
            "Nro días trabajados": v("Nro días trabajados en el mes(1115)"),
            "Nro días de licencia médica": v("Nro días de licencia médica en el mes(1116)"),
            "Sueldo(2101)": v("Sueldo(2101)"), "Sobresueldo(2102)": v("Sobresueldo(2102)"),
            "Comisiones(2103)": v("Comisiones(2103)"), "Semana corrida(2104)": v("Semana corrida(2104)"),
            "Participación(2105)": v("Participación(2105)"), "Gratificación(2106)": v("Gratificación(2106)"),
            "Recargo 30% día domingo (Art. 38) (2107)": v("Recargo 30% día domingo(2107)"),
            "Remuneración variable pagada en vacaciones (Art 71) (cód 2108)": v("Remun. variable pagada en vacaciones(2108)"),
            "Aguinaldo(2110)": v("Aguinaldo(2110)"),
            "Bonos u otras remun. fijas mensuales(2111)": v("Bonos u otras remun. fijas mensuales(2111)"),
            "Tratos (mensual) (cód 2112)": v("Tratos(2112)"),
            "Bonos u otras remuneraciones variables mensuales o superiores a un mes (cód 2113)": v("Bonos u otras remun. variables mensuales o superiores a un mes(2113)"),
            "Beneficios en especie constitutivos de remuneración (cód 2115)": v("Beneficios en especie constitutivos de remun(2115)"),
            "Otras remuneraciones superiores a un mes (cód 2123)": v("Otras remuneraciones superiores a un mes(2123)"),
            "Pago por horas de trabajo sindical (cód 2124)": v("Pago por horas de trabajo sindical(2124)"),
            "Subsidio por incapacidad laboral por licencia médica(2201)": v("Subsidio por incapacidad laboral por licencia médica(2201)"),
            "Beca de estudio (Art. 17 N°18 LIR) (cód 2202)": v("Beca de estudio(2202)"),
            "Otros ingresos no constitutivos de renta (Art 17 N°29 LIR) (cód 2204)": v("Otros ingresos no constitutivos de renta(2204)"),
            "Colación(2301)": v("Colación(2301)"), "Movilización(2302)": v("Movilización(2302)"),
            "Viáticos totales mensual (Art 41) (cód 2303)": v("Viáticos(2303)"),
            "Asignación de pérdida de caja(2304)": v("Asignación de pérdida de caja(2304)"),
            "Asignación de desgaste herramienta(2305)": v("Asignación de desgaste herramienta(2305)"),
            "Gastos por causa del trabajo (Art 41 CdT) y gastos de representación (Art. 42 Nº1 LIR) (cód 2306)": v("Gastos por causa del trabajo(2306)"),
            "Sala cuna (Art 203) (cód 2308)": v("Sala cuna(2308)"),
            "Asignación familiar legal(2311)": v("Asignación familiar legal(2311)"),
            "Asignación trabajo a distancia o teletrabajo(2309)": v("Asignación trabajo a distancia o teletrabajo(2309)"),
            "Alojamiento por razones de trabajo (2310)": v("Alojamiento por razones de trabajo(2310)"),
            "Asignación de traslación(2312)": v("Asignación de traslación(2312)"),
            "Indemnización por feriado legal(2313)": v("Indemnización por feriado legal(2313)"),
            "Indemnización años de servicio(2314)": v("Indemnización años de servicio(2314)"),
            "Indemnización sustitutiva del aviso previo(2315)": v("Indemnización sustitutiva del aviso previo(2315)"),
            "Indemnización fuero maternal (Art 163 bis) (cód 2316)": v("Indemnización fuero maternal(2316)"),
            "Indemnización a todo evento (Art.164) (cód 2331)": v("Pago indemnización a todo evento(2331)"),
            "Indemnizaciones voluntarias tributables (cód 2417)": v("Indemnizaciones voluntarias tributables(2417)"),
            "Indemnizaciones contractuales tributables (cód 2418)": v("Indemnizaciones contractuales tributables(2418)"),
            "Cotización obligatoria previsional (AFP o IPS)(3141)": v("Cotización obligatoria previsional (AFP o IPS)(3141)"),
            "Cotización obligatoria salud 7%(3143)": v("Cotización obligatoria salud 7%(3143)"),
            "Cotización voluntaria para salud(3144)": v("Cotización voluntaria para salud(3144)"),
            "Cotización AFC - trabajador(3151)": v("Cotización AFC - trabajador(3151)"),
            "Cotización adicional trabajo pesado- trabajador (cód 3154)": v("Cotización adicional trabajo pesado - trabajador(3154)"),
            "Cotización APVi Mod A(3155)": v("Cotización APVi Mod A(3155)"),
            "Cotización APVi Mod B hasta UF50(3156)": v("Cotización APVi Mod B hasta UF50(3156)"),
            "Impuesto retenido por remuneraciones(3161)": v("Impuesto retenido por remuneraciones(3161)"),
            "Impuesto retenido por indemnizaciones (cód 3162)": v("Impuesto retenido por indemnizaciones(3162)"),
            "Mayor retención de impuesto solicitada por el trabajador (cód 3163)": v("Mayor retención de impuestos solicitada por el trabajador(3163)"),
            "Impuesto retenido por reliquidación de remuneraciones devengadas en otros períodos mensuales (cód 3164)": v("Impuesto retenido por reliquidación remun. devengadas otros períodos(3164)"),
            "Retención préstamo clase media 2020 (Ley 21.252) (3166)": v("Retención préstamo clase media 2020 (Ley 21.252) (3166)"),
            "Cuota sindical 1(3171)": v("Cuota sindical 1(3171)"),
            "Crédito social CCAF(3110)": v("Crédito social CCAF(3110)"),
            "Cuota vivienda o educación Art. 58 (cód 3181)": v("Cuota vivienda o educación(3181)"),
            "Crédito cooperativas de ahorro (Art 54 Ley Coop.) (cód 3182)": v("Crédito cooperativas de ahorro(3182)"),
            "Otros descuentos autorizados y solicitados por el trabajador (cód 3183)": v("Otros descuentos autorizados y solicitados por el trabajador(3183)"),
            "Otros descuentos(3185)": v("Otros descuentos(3185)"),
            "Pensiones de alimentos(3186)": v("Pensiones de alimentos(3186)"),
            "Descuentos por anticipos y préstamos(3188)": v("Descuentos por anticipos y préstamos(3188)"),
            "AFC - Aporte empleador solidario": afc_input,
            "AFC - Aporte empleador individual": 0,
            "Aporte empleador seguro accidentes del trabajo y Ley SANNA(4152)": v("Aporte empleador seguro accidentes del trabajo y Ley SANNA(4152)"),
            "Aporte adicional trabajo pesado- empleador (cód 4154)": v("Aporte adicional trabajo pesado - empleador(4154)"),
            "Rebaja zona extrema DL 889 (3167)": v("Rebaja zona extrema DL 889 (3167)"),
            "Aporte empleador seguro invalidez y sobrevivencia(4155)": v("Aporte empleador seguro invalidez y sobrevivencia(4155)"),
            "Total líquido(5501)": v("Total líquido(5501)"),
            "jornada": jornada,
            "Validacion Liquido": val_liq,
            "AfcSolidValidacion": afc_solid_val,
            "AfcIndiValidacion": afc_indi_val,
            "DiferenciasAFC": dif_afc,
        })

    df_out = pd.DataFrame(filas)
    # Quitar cero inicial de RUTs — Rex+ no acepta formato 05829866-2
    df_out["Id empleado"] = df_out["Id empleado"].apply(
        lambda r: r.lstrip("0") if isinstance(r, str) and r.startswith("0") else r
    )
    return df_out.sort_values("Id empleado").reset_index(drop=True), rut_empresa, periodo


# ══════════════════════════════════════════════════════════════════════════════
# INTERFAZ — 4 pasos
# ══════════════════════════════════════════════════════════════════════════════

def step_header(num, label, done=False):
    bg      = "#1EBBEF" if done else "#1A3A5F"
    icon    = "✓" if done else str(num)
    suffix  = " ✓" if done else ""
    st.markdown(f"""
    <div style="display:inline-flex;align-items:center;gap:8px;
                background:white;border:1px solid #dde3f0;border-radius:99px;
                padding:5px 14px 5px 6px;margin-bottom:10px;
                box-shadow:0 1px 3px rgba(0,0,0,0.06);">
        <span style="background:{bg};color:white;border-radius:50%;
                     width:22px;height:22px;display:inline-flex;
                     align-items:center;justify-content:center;
                     font-size:11px;font-weight:700;">{num}</span>
        <span style="font-size:13px;font-weight:600;color:#1A3A5F;">{label}{suffix}</span>
    </div>""", unsafe_allow_html=True)

# ── PASO 1: Archivos de referencia del cliente ────────────────────────────────
paso1_ok = st.session_state.empleado_bytes and st.session_state.empresa_bytes
step_header(1, "Archivos de referencia del cliente", done=paso1_ok)

with st.expander("👥 empleado.xlsx · 🏢 empresa.xlsx", expanded=not paso1_ok):
    st.caption("Estos archivos se usan solo en esta sesión y no se guardan en ningún servidor.")

    # ── Instrucciones SQL ─────────────────────────────────────────────────────
    with st.expander("📋 ¿Cómo generar estos archivos desde Rex+?", expanded=False):
        st.markdown("""
        Ejecuta cada consulta en Rex+ (Módulo de reportes / consulta directa),
        exporta el resultado como **.xlsx** y súbelo aquí.
        """)
        st.markdown("**Consulta empleado.xlsx:**")
        SQL_EMPLEADO = """select contr.empleado as "Id empleado",
       contr.contrato as "Numero de contrato",
       contr."fechaInic" as "fechaInic",
       contr."tipoCont" as "tipoCont",
       (select emp.jubilado from T$empleados as emp
        where contr.empleado = emp.empleado) as "jubilado"
from T$empleadoscontr as contr"""
        st.code(SQL_EMPLEADO, language="sql")

        st.markdown("**Consulta empresa.xlsx:**")
        SQL_EMPRESA = """select identificador_nacional as "Rut empresa",
       empresa as "idempresa",
       mutual as "Mutual",
       "cotizacionMutu" as "% mutual"
from T$empresas"""
        st.code(SQL_EMPRESA, language="sql")

    # ── Archivos ejemplo ──────────────────────────────────────────────────────
    import openpyxl, io as _io, datetime as _dt

    def make_ejemplo_empleado():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Hoja 1"
        ws.append(["Id empleado", "Numero de contrato", "fechaInic", "tipoCont", "jubilado"])
        ws.append(["12345678-9", 1, _dt.date(2020, 1, 15), "I", 0])
        ws.append(["98765432-1", 1, _dt.date(2018, 6, 1),  "O", 0])
        ws.append(["11111111-1", 2, _dt.date(2022, 3, 10), "F", 0])
        buf = _io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf.read()

    def make_ejemplo_empresa():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Hoja 1"
        ws.append(["Rut empresa", "idempresa", "Mutual", "% mutual"])
        ws.append(["76016236-1", 1, "ACHS", 0.93])
        ws.append(["12345678-0", 2, "IST",  0.85])
        buf = _io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf.read()

    col_ej1, col_ej2 = st.columns(2)
    with col_ej1:
        st.download_button(
            "📥 Descargar ejemplo empleado.xlsx",
            data=make_ejemplo_empleado(),
            file_name="ejemplo_empleado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_ej2:
        st.download_button(
            "📥 Descargar ejemplo empresa.xlsx",
            data=make_ejemplo_empresa(),
            file_name="ejemplo_empresa.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()

    # ── Advertencia ───────────────────────────────────────────────────────────
    st.warning(
        "⚠️ **Importante:** Antes de subir los archivos, asegúrate de **eliminar la primera fila** "
        "de `empleado.xlsx` y `empresa.xlsx`. Esa fila contiene el nombre de la consulta y "
        "no debe incluirse en los datos."
    )

    # ── Uploaders ─────────────────────────────────────────────────────────────
    col1, col2 = st.columns(2)
    with col1:
        f_emp = st.file_uploader("👥 empleado.xlsx", type=["xlsx"], key="up_empleado")
        if f_emp:
            st.session_state.empleado_bytes = f_emp.read()
            st.session_state.refs_cargadas  = False
            st.success("✓ empleado.xlsx cargado")
    with col2:
        f_emp2 = st.file_uploader("🏢 empresa.xlsx", type=["xlsx"], key="up_empresa")
        if f_emp2:
            st.session_state.empresa_bytes = f_emp2.read()
            st.session_state.refs_cargadas = False
            st.success("✓ empresa.xlsx cargado")

    if paso1_ok:
        if st.button("🗑 Limpiar archivos de referencia"):
            st.session_state.empleado_bytes = None
            st.session_state.empresa_bytes  = None
            st.session_state.refs_cargadas  = False
            st.rerun()

st.divider()

# ── PASO 2: Equivalencias ─────────────────────────────────────────────────────
equiv_ok = st.session_state.equiv_override or EQUIVALENCIAS_BASE.exists()
step_header(2, "Equivalencias", done=equiv_ok)

with st.expander("⚙️ Gestión de equivalencias", expanded=not equiv_ok):
    tab1, tab2, tab3 = st.tabs(["📥 Descargar base", "📤 Usar actualizada", "💾 Actualizar archivo equivalencia (admin)"])

    with tab1:
        if EQUIVALENCIAS_BASE.exists():
            with open(EQUIVALENCIAS_BASE, "rb") as f:
                st.download_button(
                    "📥 Descargar equivalencias base",
                    data=f.read(),
                    file_name="equivalencias.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, type="primary",
                )
            st.caption("Descárgala, modifícala en Excel y súbela en la pestaña '📤 Usar actualizada'.")
        else:
            st.warning("No se encontró el archivo base. Sube una versión en la pestaña '📤 Usar actualizada'.")

    with tab2:
        up_equiv = st.file_uploader("Sube tu equivalencias.xlsx actualizada", type=["xlsx"], key="up_equiv")
        if up_equiv:
            st.session_state.equiv_override = up_equiv.read()
            st.session_state.refs_cargadas  = False
            st.success("✓ Equivalencias actualizadas para esta sesión.")
        if st.session_state.equiv_override:
            st.info("✓ Usando equivalencias personalizadas de esta sesión.")
            if st.button("🗑 Volver a usar la versión base"):
                st.session_state.equiv_override = None
                st.session_state.refs_cargadas  = False
                st.rerun()

    with tab3:
        st.caption("Crea un PR en GitHub para actualizar la versión base que todos comparten.")
        with st.expander("🔑 Token de GitHub", expanded="gh_token" not in st.session_state):
            tok = st.text_input("Personal Access Token", type="password", placeholder="ghp_xxxx")
            if st.button("Guardar token") and tok:
                st.session_state["gh_token"] = tok
                st.success("✓ Token guardado para esta sesión.")

        if "gh_token" in st.session_state:
            pr_file = st.file_uploader("Sube la equivalencias.xlsx actualizada", type=["xlsx"], key="equiv_pr")
            pr_desc = st.text_input("Descripción del cambio", placeholder="Ej: Actualizar porcentajes AFP 2025")
            if st.button("🚀 Crear PR con la nueva equivalencia", type="primary", use_container_width=True):
                if not pr_file or not pr_desc:
                    st.error("Sube el archivo y describe el cambio.")
                else:
                    import requests, base64
                    from datetime import datetime
                    token    = st.session_state["gh_token"]
                    owner    = "nicolasparra-rex"
                    repo     = "rex-tools"
                    headers  = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
                    # Obtener SHA del archivo actual
                    r = requests.get(f"https://api.github.com/repos/{owner}/{repo}/contents/Equivalencia/equivalencias.xlsx", headers=headers)
                    if r.status_code != 200:
                        r = requests.get(f"https://api.github.com/repos/{owner}/{repo}/contents/equivalencias.xlsx", headers=headers)
                    if r.status_code != 200:
                        st.error("No se pudo acceder al archivo base. Verifica el token.")
                    else:
                        sha_file    = r.json()["sha"]
                        file_path   = r.json()["path"]
                        branch_name = f"equiv-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
                        # SHA de main
                        r2 = requests.get(f"https://api.github.com/repos/{owner}/{repo}/git/ref/heads/main", headers=headers)
                        sha_main = r2.json()["object"]["sha"]
                        # Crear branch
                        requests.post(f"https://api.github.com/repos/{owner}/{repo}/git/refs", headers=headers,
                                      json={"ref": f"refs/heads/{branch_name}", "sha": sha_main})
                        # Subir archivo
                        content_b64 = base64.b64encode(pr_file.read()).decode()
                        r3 = requests.put(f"https://api.github.com/repos/{owner}/{repo}/contents/{file_path}",
                                          headers=headers,
                                          json={"message": f"Actualizar equivalencias: {pr_desc}", "content": content_b64, "sha": sha_file, "branch": branch_name})
                        # Crear PR
                        r4 = requests.post(f"https://api.github.com/repos/{owner}/{repo}/pulls", headers=headers,
                                           json={"title": f"Equivalencias: {pr_desc}", "head": branch_name, "base": "main", "body": f"Cambio: {pr_desc}"})
                        if r4.status_code in (200, 201):
                            st.success("✓ PR creado.")
                            st.markdown(f"[Ver PR →]({r4.json().get('html_url')})")
                        else:
                            st.error("No se pudo crear el PR.")

st.divider()

# ── PASO 3: Subir CSV(s) del LRE ─────────────────────────────────────────────
paso3_ok = paso1_ok and equiv_ok
step_header(3, "Subir archivo(s) LRE", done=False)

if not paso3_ok:
    st.info("Completa los pasos 1 y 2 antes de subir el CSV.")
else:
    st.markdown(
        '<div style="background:#1A3A5F;border-radius:10px;padding:1rem 1.5rem;margin-bottom:1rem;">'
        '<div style="color:#1EBBEF;font-size:0.75rem;font-weight:700;margin-bottom:0.5rem;">INSTRUCCIONES</div>'
        '<div style="color:white;font-size:0.9rem;line-height:1.8;">'
        '📂 <b>Formato:</b> CSV separado por punto y coma (;), encoding latin-1<br>'
        '📄 <b>Nombre:</b> debe terminar en <code>RUTempresa_YYYYMM.csv</code> · Ej: <code>760162361_202501.csv</code>'
        '</div></div>', unsafe_allow_html=True
    )

    archivos = st.file_uploader(
        "Selecciona uno o más archivos CSV",
        type=["csv"], accept_multiple_files=True, label_visibility="collapsed",
    )

    if archivos:
        # Determinar bytes de equivalencias
        equiv_bytes = st.session_state.equiv_override
        if not equiv_bytes and EQUIVALENCIAS_BASE.exists():
            with open(EQUIVALENCIAS_BASE, "rb") as f:
                equiv_bytes = f.read()

        with st.spinner("Cargando tablas de referencia..."):
            try:
                refs = cargar_referencias(
                    equiv_bytes,
                    st.session_state.empleado_bytes,
                    st.session_state.empresa_bytes,
                )
            except Exception as e:
                st.error(f"❌ Error cargando referencias: {e}")
                st.stop()

        st.divider()
        step_header(4, "Resultados", done=False)

        resultados, errores_proceso = [], []
        with st.spinner(f"Procesando {len(archivos)} archivo(s)..."):
            for archivo in archivos:
                try:
                    df_out, rut_emp, periodo = transformar(archivo.read(), archivo.name, refs)
                    resultados.append({"nombre": archivo.name, "rut": rut_emp, "periodo": periodo, "df": df_out})
                except Exception as e:
                    errores_proceso.append({"nombre": archivo.name, "error": str(e)})

        for ep in errores_proceso:
            st.error(f"❌ **{ep['nombre']}**: {ep['error']}")

        if resultados:
            total_filas   = sum(len(r["df"]) for r in resultados)
            total_revisar = sum((r["df"]["Número de contrato"] == "revisar").sum() for r in resultados)
            total_dif_afc = sum((r["df"]["DiferenciasAFC"] == "revisar con cliente").sum() for r in resultados)
            total_liq_dif = sum((r["df"]["Validacion Liquido"] != 0).sum() for r in resultados)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("📄 Archivos", len(resultados))
            c2.metric("👤 Trabajadores", total_filas)
            c3.metric("⚠️ Sin tipo contrato", total_revisar)
            c4.metric("🔍 AFC a revisar", total_dif_afc)
            if total_liq_dif > 0:
                st.warning(f"⚠️ {total_liq_dif} trabajador(es) con diferencia en líquido.")

            for r in resultados:
                df = r["df"]
                with st.expander(f"📄 {r['nombre']} — {r['rut']} | {r['periodo']} | {len(df)} trabajadores", expanded=len(resultados) == 1):
                    sc = df[df["Número de contrato"] == "revisar"]
                    if not sc.empty:
                        st.warning(f"⚠️ Sin tipo contrato: {', '.join(sc['Id empleado'].tolist())}")
                    ar = df[df["DiferenciasAFC"] == "revisar con cliente"]
                    if not ar.empty:
                        st.warning(f"⚠️ AFC a revisar: {', '.join(ar['Id empleado'].tolist())}")
                    lr = df[df["Validacion Liquido"] != 0]
                    if not lr.empty:
                        st.warning(f"⚠️ Diferencia en líquido: {', '.join(lr['Id empleado'].tolist())}")
                    if sc.empty and ar.empty and lr.empty:
                        st.success("✅ Sin observaciones.")

                    st.dataframe(
                        df[["Id empleado","Número de contrato","afp","isapre",
                            "Nro días trabajados","Total líquido(5501)",
                            "Validacion Liquido","DiferenciasAFC"]],
                        use_container_width=True, hide_index=True,
                    )

                    buf = io.BytesIO()
                    buf.write("\ufeff".encode("utf-8"))
                    buf.write(df.to_csv(index=False, encoding="utf-8", decimal=",").encode("utf-8"))
                    buf.seek(0)
                    st.download_button(
                        f"📥 Descargar {os.path.splitext(r['nombre'])[0]}_salida.csv",
                        data=buf, file_name=f"{os.path.splitext(r['nombre'])[0]}_salida.csv",
                        mime="text/csv", type="primary", key=f"dl_{r['nombre']}",
                    )

            if len(resultados) > 1:
                st.markdown("---")
                df_total  = pd.concat([r["df"] for r in resultados], ignore_index=True)
                buf_total = io.BytesIO()
                buf_total.write("\ufeff".encode("utf-8"))
                buf_total.write(df_total.to_csv(index=False, encoding="utf-8", decimal=",").encode("utf-8"))
                buf_total.seek(0)
                st.download_button(
                    f"📥 Descargar consolidado ({len(df_total)} trabajadores)",
                    data=buf_total, file_name="lre_consolidado.csv",
                    mime="text/csv", type="primary", key="dl_consolidado",
                )

aplicar_footer()
