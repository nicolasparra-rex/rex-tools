# -*- coding: utf-8 -*-
"""Rex-tools · Migración detalle desde el Libro de Remuneraciones de cualquier cliente.
Lee un libro con estructura arbitraria, autodetecta período/estructura, propone el mapeo de
conceptos contra el catálogo del cliente y genera la planilla de migración detalle, cuadrada
al peso. Empresa/mutual/contrato/caja se resuelven por RUT desde la dotación."""
import streamlit as st
import pandas as pd
import io, os, json
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

from libro_engine import (norm, load_grid, detect_header_row, match_struct, classify_and_map,
                          generar_detalle, cargar_homologacion, cargar_dotacion, detectar_periodo,
                          OUT_COLS)

try:
    from lib.branding import aplicar_branding, aplicar_footer, hero
except Exception:
    def aplicar_branding(**k): pass
    def aplicar_footer(): pass
    def hero(t, d="", i=""): st.title(t); st.caption(d)

DATA_DIR = "data"
AMBAR = "#FFF3CD"; ROJO = "#F8D7DA"; VERDE = "#D4EDDA"

st.set_page_config(page_title="Rex+ | Migración desde Libro del Cliente", page_icon="📘", layout="wide")
aplicar_branding(titulo_pagina="Migración desde Libro", badge="BETA")
hero("📘 Migración detalle desde el Libro del Cliente",
     "Sube el libro del cliente y la dotación, confirma el mapeo de conceptos y genera la migración detalle cuadrada al peso.")

# ---------- helpers de datos de referencia ----------
@st.cache_data(show_spinner=False)
def cargar_parametros():
    wp = load_workbook(os.path.join(DATA_DIR, "parametrosMesuales.xlsx"), data_only=True).worksheets[0]
    ph = [wp.cell(row=1, column=c).value for c in range(1, wp.max_column + 1)]
    return {wp.cell(row=r, column=1).value: {ph[c-1]: wp.cell(row=r, column=c).value
            for c in range(1, wp.max_column + 1)} for r in range(2, wp.max_row + 1) if wp.cell(row=r, column=1).value}

@st.cache_data(show_spinner=False)
def cargar_cot_hist():
    ch = load_workbook(os.path.join(DATA_DIR, "cot_afp_hist.xlsx"), data_only=True).worksheets[0]
    return {f"{ch.cell(row=r,column=2).value}{ch.cell(row=r,column=3).value}": (ch.cell(row=r,column=5).value or 0)
            for r in range(2, ch.max_row + 1)}

@st.cache_data(show_spinner=False)
def cargar_homolog_default():
    p = os.path.join(DATA_DIR, "listado_instituciones.xlsx")
    return cargar_homologacion(p) if os.path.exists(p) else []

def leer_catalogo(file):
    df = pd.read_excel(file, header=None)
    hr = 0
    for r in range(min(6, len(df))):
        vals = [norm(x) for x in df.iloc[r].values]
        if "concepto" in vals and any("nombre" in v for v in vals): hr = r; break
    hdr = [norm(x) for x in df.iloc[hr].values]
    ci = hdr.index("concepto") if "concepto" in hdr else 0
    ni = next((i for i, v in enumerate(hdr) if "nombre" in v), 2)
    names, valid = {}, {}
    for _, row in df.iloc[hr+1:].iterrows():
        cid = row[ci]; nom = row[ni]
        if pd.notna(cid) and str(cid).strip():
            valid[str(cid).strip()] = (str(nom).strip() if pd.notna(nom) else "")
            if pd.notna(nom): names[norm(nom)] = str(cid).strip()
    return names, valid

def to_excel(filas):
    wb = Workbook(); o = wb.active; o.title = "Migración"
    hf = PatternFill("solid", fgColor="1A2744"); ff = Font(color="FFFFFF", bold=True, size=10)
    for j, cn in enumerate(OUT_COLS, 1):
        c = o.cell(row=1, column=j, value=cn); c.fill = hf; c.font = ff
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, rr in enumerate(filas, 2):
        for j, v in enumerate(rr, 1): o.cell(row=i, column=j, value=v)
    o.freeze_panes = "A2"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

# ---------- Barra lateral: solo lo mínimo ----------
with st.sidebar:
    st.subheader("⚙️ Datos del cliente")
    cliente = st.text_input("Nombre del cliente", value="", placeholder="ej. thoughtworks")
    st.caption("Empresa, mutual, contrato y caja se resuelven por RUT desde la **dotación**. "
               "El período se detecta del archivo.")
    with st.expander("Avanzado (valores por defecto)"):
        apv_inst = st.text_input("Institución APV / Cuenta 2", value="afp")
        caja_inst = st.text_input("Caja CCAF (si la dotación no la trae)", value="losandes")
        jornada = st.text_input("Jornada", value="C")
# fallbacks internos (se completan desde la dotación)
empresa_id, mutual_id, num_contrato = "", "", 1

# ---------- 1. Archivos ----------
st.markdown("### 1 · Sube los archivos")
c1, c2 = st.columns(2)
libro_file = c1.file_uploader("① Libro de remuneraciones del cliente", type=["xlsx", "xls"])
cat_file = c2.file_uploader("② Catálogo de conceptos del cliente (necesario para elegir los IDs)", type=["xlsx"])
c3, c4 = st.columns(2)
dot_file = c3.file_uploader("③ Dotación (RUT → contrato / empresa / mutual / caja)", type=["xlsx"])
map_file = c4.file_uploader("④ Mapeo guardado del cliente (opcional, .json)", type=["json"])
with st.expander("Homologación de instituciones (avanzado)"):
    homolog_file = st.file_uploader("Reemplazar homologación (opcional; por defecto usa data/listado_instituciones.xlsx)", type=["xlsx"])
    _hp = os.path.join(DATA_DIR, "listado_instituciones.xlsx")
    if os.path.exists(_hp):
        with open(_hp, "rb") as _f:
            st.download_button("⬇️ Descargar homologación actual", _f.read(), file_name="listado_instituciones.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.caption("Edítala y vuelve a subirla para actualizar la homologación.")

if not libro_file:
    st.info("Sube el **libro** y el **catálogo** para comenzar. La **dotación** completa empresa/mutual/contrato/caja.")
    aplicar_footer(); st.stop()

catalog_names, valid_map = ({}, {})
if cat_file: catalog_names, valid_map = leer_catalogo(cat_file)
valid_ids = sorted(valid_map.keys())
saved = {}
if map_file: saved = {norm(k): v for k, v in json.load(map_file).items()}
homolog = cargar_homologacion(homolog_file) if homolog_file else cargar_homolog_default()
dotacion = cargar_dotacion(dot_file) if dot_file else {}

df, sheet = load_grid(libro_file)
hr = detect_header_row(df)
hdr = [x if str(x) != "nan" else "" for x in df.iloc[hr].values]
struct = match_struct(hdr)
propuesta = classify_and_map(hdr, struct, catalog_names=catalog_names, saved=saved)

# ---------- 2. Período + estructura ----------
st.markdown("### 2 · Período y estructura detectados")
periodo = detectar_periodo(df, libro_file.name)
ca, cb = st.columns([1, 3])
if periodo:
    ca.success(f"📅 Período: **{periodo}**")
else:
    periodo = ca.text_input("📅 No pude detectar el período — escríbelo (AAAA-MM)", value="", placeholder="2026-06")
cb.caption(f"Hoja: **{sheet}** · Encabezado en fila **{hr+1}**")

faltan = [k for k in ["rut", "total_haberes", "total_descuentos", "liquido", "base_afp"] if k not in struct]
if faltan:
    st.error(f"❌ No detecté columnas estructurales: **{', '.join(faltan)}**. Revisa el libro antes de continuar.")
if not dotacion:
    st.warning("⚠️ No subiste la **dotación**: empresa, mutual, contrato y caja no se podrán resolver por RUT.")

# ---------- 3. Mapeo ----------
st.markdown("### 3 · Confirma el mapeo de conceptos")
if not valid_ids:
    st.error("❌ Sube el **catálogo de conceptos del cliente** para poder elegir los IDs de una lista.")
    aplicar_footer(); st.stop()

st.caption("Elige el **ID Rex** de cada concepto en el desplegable. En **ámbar** los que faltan por mapear. "
           "Cuando esté todo, se habilita el botón de generar. Al final puedes descargar el mapeo para reutilizarlo.")
map_df = pd.DataFrame(propuesta)[["col", "header", "grupo", "id_rex", "fuente", "confianza"]]
tot = len(map_df); auto0 = int((map_df["id_rex"].astype(str).str.len() > 0).sum())
m1, m2, m3 = st.columns(3)
m1.metric("Conceptos", tot); m2.metric("Sugeridos", auto0); m3.metric("Por confirmar", tot - auto0)

editor = st.data_editor(
    map_df, use_container_width=True, hide_index=True, key="mapeo", height=430,
    column_config={
        "col": st.column_config.NumberColumn("Col", disabled=True, width="small"),
        "header": st.column_config.TextColumn("Columna del libro", disabled=True),
        "grupo": st.column_config.TextColumn("Bloque", disabled=True, width="small"),
        "id_rex": st.column_config.SelectboxColumn("ID Rex (elige)", options=valid_ids, required=False,
                    help="Concepto Rex al que corresponde esta columna del libro."),
        "fuente": st.column_config.TextColumn("Fuente", disabled=True, width="small"),
        "confianza": st.column_config.TextColumn("Conf.", disabled=True, width="small"),
    })

mapping = {norm(r["header"]): r["id_rex"] for _, r in editor.iterrows() if r["id_rex"]}
invalidos = [v for v in set(mapping.values()) if v not in valid_map]
pend_df = editor[editor["id_rex"].isna() | (editor["id_rex"].astype(str).str.strip() == "")][["col", "header", "grupo"]]

if len(pend_df):
    st.warning(f"🟠 Faltan **{len(pend_df)}** concepto(s) por mapear — complétalos arriba:")
    st.dataframe(pend_df.rename(columns={"header": "Columna del libro", "grupo": "Bloque"})
                 .style.set_properties(**{"background-color": AMBAR}), hide_index=True, use_container_width=True)
else:
    st.success("✅ Todos los conceptos están mapeados.")
if invalidos:
    st.error(f"❌ IDs que no existen en el catálogo: {invalidos}")

# ---------- 4. Generar ----------
st.markdown("### 4 · Generar y validar")
listo = bool(periodo) and not faltan and not len(pend_df) and not invalidos
if not listo:
    st.info("Para habilitar: período detectado, estructura completa, **0 conceptos por mapear** y sin IDs inválidos.")

cgen, cmap = st.columns([1, 1])
if cgen.button("🚀 Generar migración detalle", type="primary", disabled=not listo, use_container_width=True):
    cfg = dict(empresa_id=empresa_id, mutual_id=mutual_id, apv_inst=apv_inst, caja_inst=caja_inst,
               num_contrato=int(num_contrato), jornada=jornada, periodo=periodo)
    filas, res = generar_detalle(df, hr, struct, mapping, cargar_parametros().get(periodo, {}),
                                 cargar_cot_hist(), cfg, homolog=homolog, dotacion=dotacion)
    ok = (res["descuadre_haberes"] == 0 and res["descuadre_descuentos"] == 0 and res["descuadre_liquido"] == 0)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Empleados", res["empleados"]); k2.metric("Filas", len(filas))
    k3.metric("Descuadres", res["descuadre_haberes"] + res["descuadre_descuentos"] + res["descuadre_liquido"])
    k4.metric("Estado", "✅ Cuadra" if ok else "⚠️ Revisar")
    if not ok:
        st.error(f"Descuadres → haberes {res['descuadre_haberes']}, descuentos {res['descuadre_descuentos']}, "
                 f"líquido {res['descuadre_liquido']}.")
    if res.get("log_contratos"):
        lc = pd.DataFrame(res["log_contratos"]).rename(columns={"rut": "RUT", "motivo": "Motivo"})
        st.error(f"🔴 {len(lc)} RUT con problema de contrato/dotación (revisar):")
        st.dataframe(lc.style.set_properties(**{"background-color": ROJO}), hide_index=True, use_container_width=True)
        st.download_button("⬇️ Log rut-contrato (.csv)", lc.to_csv(index=False).encode("utf-8"),
                           file_name=f"log_contratos_{cliente or 'cliente'}_{periodo}.csv", mime="text/csv")
    for f in res["flags"]:
        st.warning("🟠 " + f)
    st.download_button("⬇️ Descargar migración detalle (.xlsx)", to_excel(filas),
                       file_name=f"migracion_detalle_{cliente or 'cliente'}_{periodo}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

cmap.download_button("💾 Descargar mapeo del cliente (.json) para reutilizar",
                     json.dumps({r["header"]: r["id_rex"] for _, r in editor.iterrows() if r["id_rex"]},
                                ensure_ascii=False, indent=2),
                     file_name=f"mapeo_{cliente or 'cliente'}.json", mime="application/json", use_container_width=True)

aplicar_footer()
