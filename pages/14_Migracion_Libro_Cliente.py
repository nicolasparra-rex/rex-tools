# -*- coding: utf-8 -*-
"""Rex-tools · Migración detalle desde el Libro de Remuneraciones de cualquier cliente.
Lee un libro con estructura arbitraria, autodetecta columnas, propone el mapeo de conceptos
contra el catálogo del cliente y genera la planilla de migración detalle, validada al peso."""
import streamlit as st
import pandas as pd
import io, os, json, unicodedata
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

from libro_engine import (norm, load_grid, detect_header_row, match_struct,
                          classify_and_map, generar_detalle, cargar_homologacion, cargar_dotacion,
                          OUT_COLS, STRUCT)

try:
    from lib.branding import aplicar_branding, aplicar_footer, hero
except Exception:
    def aplicar_branding(**k): pass
    def aplicar_footer(): pass
    def hero(t, d="", i=""): st.title(t); st.caption(d)

DATA_DIR = "data"
st.set_page_config(page_title="Rex+ | Migración desde Libro del Cliente", page_icon="📘", layout="wide")
aplicar_branding(titulo_pagina="Migración desde Libro", badge="BETA")
hero("📘 Migración detalle desde el Libro del Cliente",
     "Sube el libro de remuneraciones de cualquier cliente, confirma el mapeo y genera la planilla de migración detalle, cuadrada al peso.")

# ---------- helpers de datos de referencia ----------
@st.cache_data(show_spinner=False)
def cargar_parametros():
    p = os.path.join(DATA_DIR, "parametrosMesuales.xlsx")
    wp = load_workbook(p, data_only=True).worksheets[0]
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
    """Detecta col de id (Concepto) y de Nombre. Devuelve (catalog_names, valid_ids)."""
    df = pd.read_excel(file, header=None)
    hr = 0
    for r in range(min(6, len(df))):
        vals = [norm(x) for x in df.iloc[r].values]
        if "concepto" in vals and any("nombre" in v for v in vals): hr = r; break
    hdr = [norm(x) for x in df.iloc[hr].values]
    ci = hdr.index("concepto") if "concepto" in hdr else 0
    ni = next((i for i, v in enumerate(hdr) if "nombre" in v), 2)
    names, valid = {}, set()
    for _, row in df.iloc[hr+1:].iterrows():
        cid = row[ci]; nom = row[ni]
        if pd.notna(cid) and str(cid).strip():
            valid.add(str(cid).strip());
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

# ---------- 1. Configuración del cliente ----------
with st.sidebar:
    st.subheader("⚙️ Configuración del cliente")
    cliente = st.text_input("Cliente", value="", placeholder="ej. thoughtworks")
    periodo = st.text_input("Período (AAAA-MM)", value="", placeholder="2026-06")
    empresa_id = st.text_input("ID Empresa en Rex", value="")
    mutual_id = st.text_input("Institución mutual", value="achs")
    apv_inst = st.text_input("Institución APV / Cuenta 2", value="afp")
    caja_inst = st.text_input("Institución caja (CCAF)", value="losandes")
    num_contrato = st.number_input("Número de contrato", 1, 9, 1)
    jornada = st.text_input("Jornada", value="C")

st.markdown("### 1 · Sube los archivos")
c1, c2 = st.columns(2)
libro_file = c1.file_uploader("Libro de remuneraciones del cliente", type=["xlsx", "xls"])
cat_file = c2.file_uploader("Catálogo de conceptos del cliente (opcional pero recomendado)", type=["xlsx"])
map_file = st.file_uploader("Mapeo guardado del cliente (opcional, .json)", type=["json"])
homolog_file = st.file_uploader("Homologación de instituciones (opcional; por defecto usa data/listado_instituciones.xlsx)", type=["xlsx"])
dot_file = st.file_uploader("Dotación (RUT → contrato / empresa / mutual)", type=["xlsx"])

# --- bajar / subir la tabla de homologación de data/ ---
_homolog_path = os.path.join(DATA_DIR, "listado_instituciones.xlsx")
if os.path.exists(_homolog_path):
    with open(_homolog_path, "rb") as _f:
        st.download_button("⬇️ Descargar homologación actual (para actualizarla)", _f.read(),
                           file_name="listado_instituciones.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("Edita esa tabla y vuelve a subirla arriba para actualizar la homologación de instituciones.")

if not libro_file:
    st.info("Sube el libro para comenzar. Con el catálogo del cliente el auto-mapeo es mucho mayor.")
    aplicar_footer(); st.stop()

# ---------- 2. Detección + mapeo ----------
catalog_names, valid_ids = ({}, set())
if cat_file: catalog_names, valid_ids = leer_catalogo(cat_file)
saved = {}
if map_file: saved = {norm(k): v for k, v in json.load(map_file).items()}
homolog = cargar_homologacion(homolog_file) if homolog_file else cargar_homolog_default()
dotacion = cargar_dotacion(dot_file) if dot_file else {}

df, sheet = load_grid(libro_file)
hr = detect_header_row(df)
hdr = [x if str(x) != "nan" else "" for x in df.iloc[hr].values]
struct = match_struct(hdr)
propuesta = classify_and_map(hdr, struct, catalog_names=catalog_names, saved=saved)

st.markdown("### 2 · Estructura detectada")
faltan = [k for k in ["rut", "total_haberes", "total_descuentos", "liquido", "base_afp"] if k not in struct]
cols = st.columns(4)
for idx, (campo, i) in enumerate(sorted(struct.items())):
    cols[idx % 4].metric(campo, hdr[i] if i is not None else "—")
if faltan:
    st.warning(f"⚠️ No detecté estas columnas estructurales: **{', '.join(faltan)}**. "
               "Ajusta el libro o mapéalas manualmente antes de generar.")
st.caption(f"Hoja: **{sheet}** · Fila de encabezado: **{hr+1}**")

st.markdown("### 3 · Confirma el mapeo de conceptos")
map_df = pd.DataFrame(propuesta)
sin = map_df[map_df["id_rex"].isna() | (map_df["id_rex"] == "")]
st.write(f"Conceptos: **{len(map_df)}** · auto-mapeados: **{len(map_df)-len(sin)}** · "
         f"por confirmar: **{len(sin)}**")
opciones = sorted(valid_ids) if valid_ids else sorted(set(x for x in map_df["id_rex"].dropna()))
editor = st.data_editor(
    map_df, use_container_width=True, hide_index=True, key="mapeo",
    column_config={
        "col": st.column_config.NumberColumn("Col", disabled=True, width="small"),
        "header": st.column_config.TextColumn("Columna del libro", disabled=True),
        "grupo": st.column_config.TextColumn("Bloque", disabled=True, width="small"),
        "id_rex": (st.column_config.SelectboxColumn("ID Rex", options=opciones, required=False)
                   if opciones else st.column_config.TextColumn("ID Rex")),
        "fuente": st.column_config.TextColumn("Fuente", disabled=True, width="small"),
        "confianza": st.column_config.TextColumn("Conf.", disabled=True, width="small"),
    })

# validar ids contra catálogo
mapping = {norm(r["header"]): r["id_rex"] for _, r in editor.iterrows() if r["id_rex"]}
invalidos = [v for v in set(mapping.values()) if valid_ids and v not in valid_ids]
pendientes = [r["header"] for _, r in editor.iterrows() if not r["id_rex"]]
if invalidos: st.error(f"IDs que no existen en el catálogo: {invalidos}")
if pendientes: st.warning(f"Quedan {len(pendientes)} concepto(s) sin mapear: {pendientes}")

# ---------- 4. Generar ----------
st.markdown("### 4 · Generar y validar")
listo = periodo and empresa_id and not invalidos and not faltan
if not listo:
    st.info("Completa período, ID empresa, resuelve IDs inválidos y columnas estructurales faltantes para habilitar.")
if st.button("🚀 Generar migración detalle", type="primary", disabled=not listo):
    cfg = dict(empresa_id=int(empresa_id) if str(empresa_id).isdigit() else empresa_id,
               mutual_id=mutual_id, apv_inst=apv_inst, caja_inst=caja_inst,
               num_contrato=int(num_contrato), jornada=jornada, periodo=periodo)
    params_row = cargar_parametros().get(periodo, {})
    if not params_row: st.warning(f"No hay parámetros para {periodo} en data/. Topes/SIS irán en 0.")
    filas, res = generar_detalle(df, hr, struct, mapping, params_row, cargar_cot_hist(), cfg, homolog=homolog, dotacion=dotacion)
    ok = (res["descuadre_haberes"] == 0 and res["descuadre_descuentos"] == 0 and res["descuadre_liquido"] == 0)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Empleados", res["empleados"]); m2.metric("Filas", len(filas))
    m3.metric("Descuadres", res["descuadre_haberes"] + res["descuadre_descuentos"] + res["descuadre_liquido"])
    m4.metric("Estado", "✅ Cuadra" if ok else "⚠️ Revisar")
    if not ok:
        st.error(f"Descuadres → haberes: {res['descuadre_haberes']}, descuentos: {res['descuadre_descuentos']}, "
                 f"líquido: {res['descuadre_liquido']}. Revisa el mapeo (¿algún concepto sin mapear o mal clasificado?).")
    if res.get("log_contratos"):
        lc = pd.DataFrame(res["log_contratos"])
        st.warning(f"⚠️ {len(lc)} RUT con problema de contrato/dotación (revisar):")
        st.dataframe(lc, use_container_width=True, hide_index=True)
        st.download_button("⬇️ Descargar log rut-contrato (.csv)", lc.to_csv(index=False).encode("utf-8"),
                           file_name=f"log_contratos_{cliente or 'cliente'}_{periodo}.csv", mime="text/csv")
    for f in res["flags"]: st.warning("⚠️ " + f)
    st.download_button("⬇️ Descargar migración detalle (.xlsx)", to_excel(filas),
                       file_name=f"migracion_detalle_{cliente or 'cliente'}_{periodo}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("💾 Guardar mapeo del cliente (.json)",
                       json.dumps({r["header"]: r["id_rex"] for _, r in editor.iterrows() if r["id_rex"]},
                                  ensure_ascii=False, indent=2),
                       file_name=f"mapeo_{cliente or 'cliente'}.json", mime="application/json")

aplicar_footer()
