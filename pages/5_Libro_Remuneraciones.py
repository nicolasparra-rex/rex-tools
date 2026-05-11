h"""
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
    page_title="Libro de Remuneraciones | Rex+ Tools",
    page_icon="📊",
    layout="wide",
)

# ── Branding ──────────────────────────────────────────────────────────────────
aplicar_branding(titulo_pagina="Libro de Remuneraciones")

hero(
    titulo="Libro de Remuneraciones Electrónico",
    descripcion="Sube el CSV del LRE para transformarlo al formato de importación Rex+. El nombre del archivo debe seguir el formato: RUTempresa_YYYYMM.csv  (Ej: 760162361_202501.csv)",
    icono="📊",
)

# ── Rutas de equivalencias ────────────────────────────────────────────────────
BASE_DIR      = _ROOT
EQUIVALENCIAS = BASE_DIR / "Equivalencia" / "equivalencias.xlsx"
EMPLEADO_XL   = BASE_DIR / "Equivalencia" / "empleado.xlsx"
EMPRESA_XL    = BASE_DIR / "Equivalencia" / "empresa.xlsx"


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

def safe_get(d: dict, key, default=None):
    return d.get(key, d.get(str(key), default))

def round_awz(x):
    """Redondeo AwayFromZero — igual que Power BI."""
    return math.floor(x + 0.5) if x >= 0 else math.ceil(x - 0.5)


# ── Carga de referencias ──────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def cargar_referencias(equiv_bytes: bytes | None = None):
    """Carga las tablas de referencia. Si se pasa equiv_bytes, usa esa versión."""
    if equiv_bytes:
        xl_eq = pd.ExcelFile(io.BytesIO(equiv_bytes))
    else:
        xl_eq = pd.ExcelFile(EQUIVALENCIAS)

    df_afp     = pd.read_excel(xl_eq, "afp")
    df_porafp  = pd.read_excel(xl_eq, "porafp")
    df_salud   = pd.read_excel(xl_eq, "salud")
    df_sis     = pd.read_excel(xl_eq, "sis")
    df_ccaf    = pd.read_excel(xl_eq, "ccaf")
    df_porccaf = pd.read_excel(xl_eq, "porccaf")
    df_tijor   = pd.read_excel(xl_eq, "tijor")

    xl_emp      = pd.ExcelFile(EMPLEADO_XL)
    df_empleado = pd.read_excel(xl_emp, "Hoja 1")
    df_empresa  = pd.read_excel(EMPRESA_XL, "Hoja 1")

    return {
        "afp": df_afp, "porafp": df_porafp, "salud": df_salud,
        "sis": df_sis, "ccaf": df_ccaf, "porccaf": df_porccaf,
        "tijor": df_tijor, "empleado": df_empleado, "empresa": df_empresa,
    }


# ── Transformación ────────────────────────────────────────────────────────────
def transformar(input_bytes: bytes, nombre: str, refs: dict) -> pd.DataFrame:
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

    df_sis     = refs["sis"]
    sis_row    = df_sis[df_sis["proceso"] == periodo]
    pct_sis    = float(sis_row["porcentaje"].iloc[0]) if not sis_row.empty else 0.0

    df_porccaf  = refs["porccaf"]
    porccaf_row = df_porccaf[df_porccaf["periodo"] == periodo]
    pct_ccaf    = float(porccaf_row["porcentaje"].iloc[0]) if not porccaf_row.empty else 0.0

    df_porafp      = refs["porafp"]
    porafp_periodo = (
        df_porafp[df_porafp["proceso"] == periodo]
        .set_index("nombre")["porcentaje"].to_dict()
    )

    afp_map   = dict(zip(refs["afp"]["afp"], refs["afp"]["idrex"]))
    salud_map = {int(r["id"]): r["idrex"] for _, r in refs["salud"].iterrows() if pd.notna(r["id"])}
    ccaf_map  = {int(r["id"]): r["idrex"] for _, r in refs["ccaf"].iterrows() if pd.notna(r["id"])}
    tijor_map = dict(zip(refs["tijor"]["id"], refs["tijor"]["idrex"]))

    df_emp = refs["empleado"].copy()
    df_emp["Id empleado"] = df_emp["Id empleado"].str.upper()
    df_emp["_key"] = df_emp["Id empleado"] + "_" + \
                     df_emp["fechaInic"].dt.strftime("%Y-%m-%d").fillna("")
    df_emp_idx = df_emp.set_index("_key")

    df_in = pd.read_csv(io.BytesIO(input_bytes), encoding="latin-1", sep=";")
    df_in["Rut trabajador(1101)"] = df_in["Rut trabajador(1101)"].str.upper()

    filas = []
    for _, row in df_in.iterrows():
        rut_trab = normalizar_rut(row["Rut trabajador(1101)"])

        afp_code = int(row.get("AFP(1141)", 100))
        afp_name = safe_get(afp_map, afp_code, "afp")
        pct_afp  = float(porafp_periodo.get(afp_name, 0) or 0)

        salud_code = int(row.get("FONASA - ISAPRE(1143)", 102))
        isapre     = safe_get(salud_map, salud_code, "fonasa")

        ccaf_code = row.get("CCAF(1110)", 0)
        try:
            ccaf_name = safe_get(ccaf_map, int(ccaf_code), "sincaja")
        except (ValueError, TypeError):
            ccaf_name = "sincaja"

        jorn_code = row.get("Código tipo de jornada(1107)", 101)
        try:
            jornada = safe_get(tijor_map, int(jorn_code), "C")
        except (ValueError, TypeError):
            jornada = "C"

        fecha_inic_raw = row.get("Fecha inicio contrato(1102)", None)
        try:
            fecha_inic_lre = pd.to_datetime(fecha_inic_raw, dayfirst=True)
            fecha_inic_str = fecha_inic_lre.strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            fecha_inic_lre = None
            fecha_inic_str = ""

        emp_key  = f"{rut_trab}_{fecha_inic_str}"
        emp_info = df_emp_idx.loc[emp_key] if emp_key in df_emp_idx.index else None
        if emp_info is not None:
            num_contrato = int(emp_info.get("Numero de contrato", 1) or 1)
            tipo_cont    = str(emp_info.get("tipoCont", None) or "").strip()
            jubilado     = int(emp_info.get("jubilado", 0) or 0)
        else:
            num_contrato = "revisar"
            tipo_cont    = ""
            jubilado     = 0

        if fecha_inic_lre is not None:
            fecha_fin_mes = pd.Timestamp(f"{periodo}-01") + pd.offsets.MonthEnd(0)
            antiguedad = abs(
                (fecha_fin_mes.year  - fecha_inic_lre.year)  * 12 +
                (fecha_fin_mes.month - fecha_inic_lre.month)
            )
        else:
            antiguedad = 0

        afc_input      = float(row.get("AFC - Aporte empleador(4151)", 0) or 0)
        afc_solidario  = afc_input
        afc_individual = 0

        causal_raw = row.get("Causal término de contrato(1104)", None)
        try:
            causal = int(causal_raw)
        except (ValueError, TypeError):
            causal = None

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
            "Sueldo(2101)": v("Sueldo(2101)"),
            "Sobresueldo(2102)": v("Sobresueldo(2102)"),
            "Comisiones(2103)": v("Comisiones(2103)"),
            "Semana corrida(2104)": v("Semana corrida(2104)"),
            "Participación(2105)": v("Participación(2105)"),
            "Gratificación(2106)": v("Gratificación(2106)"),
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
            "Colación(2301)": v("Colación(2301)"),
            "Movilización(2302)": v("Movilización(2302)"),
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
            "AFC - Aporte empleador solidario": afc_solidario,
            "AFC - Aporte empleador individual": afc_individual,
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
    return df_out.sort_values("Id empleado").reset_index(drop=True), rut_empresa, periodo


# ── Interfaz ──────────────────────────────────────────────────────────────────
st.markdown(
    '<div style="background:#1A3A5F;border:1px solid #1A3A5F;border-radius:12px;padding:1.25rem 1.75rem;margin-bottom:1.5rem;">'
    '<div style="color:#1EBBEF;font-weight:700;font-size:0.75rem;letter-spacing:0.5px;margin-bottom:0.75rem;">INSTRUCCIONES</div>'
    '<div style="color:white;font-size:0.95rem;line-height:1.8;">'
    '📂 <strong>Formato de entrada:</strong> CSV separado por punto y coma (;), encoding latin-1<br>'
    '📄 <strong>Nombre del archivo:</strong> debe terminar en <code>RUTempresa_YYYYMM.csv</code> &nbsp;·&nbsp; Ej: <code>760162361_202501.csv</code><br>'
    '📥 <strong>Salida:</strong> CSV UTF-8 con BOM, separado por coma, listo para importar a Rex+'
    '</div>'
    '</div>',
    unsafe_allow_html=True,
)

st.markdown("### 📤 Subir archivo(s)")

archivos = st.file_uploader(
    "Selecciona uno o más archivos CSV del LRE",
    type=["csv"],
    accept_multiple_files=True,
)

if archivos:
    if not EQUIVALENCIAS.exists() or not EMPLEADO_XL.exists() or not EMPRESA_XL.exists():
        st.error("❌ No se encontraron los archivos de equivalencias en la carpeta `Equivalencia/`.")
        st.stop()

    # Usar equivalencias override si el usuario subió una versión actualizada
    equiv_bytes = st.session_state.get("equiv_override", None)

    with st.spinner("Cargando tablas de referencia..."):
        try:
            refs = cargar_referencias(equiv_bytes)
        except Exception as e:
            st.error(f"❌ Error cargando equivalencias: {e}")
            st.stop()

    resultados      = []
    errores_proceso = []

    with st.spinner(f"Procesando {len(archivos)} archivo(s)..."):
        for archivo in archivos:
            try:
                df_out, rut_emp, periodo = transformar(archivo.read(), archivo.name, refs)
                resultados.append({"nombre": archivo.name, "rut": rut_emp, "periodo": periodo, "df": df_out})
            except Exception as e:
                errores_proceso.append({"nombre": archivo.name, "error": str(e)})

    if errores_proceso:
        for ep in errores_proceso:
            st.error(f"❌ **{ep['nombre']}**: {ep['error']}")

    if resultados:
        st.markdown("### Resumen")

        total_filas   = sum(len(r["df"]) for r in resultados)
        total_revisar = sum((r["df"]["Número de contrato"] == "revisar").sum() for r in resultados)
        total_dif_afc = sum((r["df"]["DiferenciasAFC"] == "revisar con cliente").sum() for r in resultados)
        total_liq_dif = sum((r["df"]["Validacion Liquido"] != 0).sum() for r in resultados)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📄 Archivos procesados", len(resultados))
        c2.metric("👤 Total trabajadores",  total_filas)
        c3.metric("⚠️ Sin tipo contrato",   total_revisar)
        c4.metric("🔍 AFC a revisar",       total_dif_afc)

        if total_liq_dif > 0:
            st.warning(f"⚠️ {total_liq_dif} trabajador(es) tienen diferencia en el líquido.")

        for r in resultados:
            df = r["df"]
            with st.expander(f"📄 {r['nombre']} — {r['rut']} | {r['periodo']} | {len(df)} trabajadores", expanded=len(resultados) == 1):
                sin_contrato = df[df["Número de contrato"] == "revisar"]
                if not sin_contrato.empty:
                    st.warning(f"⚠️ {len(sin_contrato)} empleado(s) sin tipo de contrato: {', '.join(sin_contrato['Id empleado'].tolist())}")

                afc_revisar = df[df["DiferenciasAFC"] == "revisar con cliente"]
                if not afc_revisar.empty:
                    st.warning(f"⚠️ {len(afc_revisar)} empleado(s) con AFC a revisar: {', '.join(afc_revisar['Id empleado'].tolist())}")

                liq_revisar = df[df["Validacion Liquido"] != 0]
                if not liq_revisar.empty:
                    st.warning(f"⚠️ {len(liq_revisar)} empleado(s) con diferencia en líquido: {', '.join(liq_revisar['Id empleado'].tolist())}")

                if sin_contrato.empty and afc_revisar.empty and liq_revisar.empty:
                    st.success("✅ Todo correcto, sin observaciones.")

                st.dataframe(
                    df[["Id empleado", "Número de contrato", "afp", "isapre",
                        "Nro días trabajados", "Total líquido(5501)",
                        "Validacion Liquido", "DiferenciasAFC"]],
                    use_container_width=True, hide_index=True,
                )

                buf = io.BytesIO()
                buf.write("\ufeff".encode("utf-8"))
                buf.write(df.to_csv(index=False, encoding="utf-8", decimal=",").encode("utf-8"))
                buf.seek(0)

                nombre_salida = os.path.splitext(r["nombre"])[0] + "_salida.csv"
                st.download_button(
                    label=f"📥 Descargar {nombre_salida}",
                    data=buf, file_name=nombre_salida, mime="text/csv",
                    type="primary", key=f"dl_{r['nombre']}",
                )

        if len(resultados) > 1:
            st.markdown("---")
            st.markdown("### 📥 Descarga consolidada")
            df_total = pd.concat([r["df"] for r in resultados], ignore_index=True)
            buf_total = io.BytesIO()
            buf_total.write("\ufeff".encode("utf-8"))
            buf_total.write(df_total.to_csv(index=False, encoding="utf-8", decimal=",").encode("utf-8"))
            buf_total.seek(0)
            st.download_button(
                label=f"📥 Descargar archivo consolidado ({len(df_total)} trabajadores)",
                data=buf_total, file_name="lre_consolidado.csv", mime="text/csv",
                type="primary", key="dl_consolidado",
            )

# ── Gestión de equivalencias ──────────────────────────────────────────────────
try:
        from equivalencias_manager import render_equivalencias_manager
except ModuleNotFoundError:
        render_equivalencias_manager = None

aplicar_footer()try:
        from equivalencias_manager import render_equivalencias_manager
except ModuleNotFoundError:
        render_equivalencias_manager = None
