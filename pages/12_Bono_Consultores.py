"""
Bono Consultores — pagina independiente y sensible.

Aislada a proposito: no importa nada de las otras paginas, tiene su propio
set_page_config y un candado de acceso. Trabaja sobre el export de Zoho
(.xlsx) que ya generas. El calculo queda estructurado para enchufar la API
de Zoho Projects mas adelante (ver seccion cargar_proyectos()).

Para soltar en rex-tools: copiar a pages/. Para ocultarlo del menu lateral,
ver nota al final del archivo.
"""

import io
import re
import unicodedata
from datetime import date

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Bono Consultores", page_icon="🔒", layout="wide")


# --------------------------------------------------------------------------
# 1. Candado de acceso (data delicada)
# --------------------------------------------------------------------------
def _check_password() -> bool:
    """Gate simple. Define BONO_PASSWORD en .streamlit/secrets.toml."""
    expected = st.secrets.get("BONO_PASSWORD")
    if expected is None:
        st.error(
            "Falta configurar la contrasena. Agrega a `.streamlit/secrets.toml`:\n\n"
            '`BONO_PASSWORD = "tu_clave"`'
        )
        st.stop()

    if st.session_state.get("bono_auth"):
        return True

    def _entered():
        if st.session_state.get("bono_pwd") == expected:
            st.session_state.bono_auth = True
            st.session_state.pop("bono_pwd", None)
        else:
            st.session_state.bono_auth = False

    st.title("🔒 Bono Consultores")
    st.caption("Acceso restringido — informacion de compensaciones.")
    st.text_input("Contrasena", type="password", key="bono_pwd", on_change=_entered)
    if st.session_state.get("bono_auth") is False:
        st.error("Contrasena incorrecta.")
    return False


if not _check_password():
    st.stop()


# --------------------------------------------------------------------------
# 2. Utilidades
# --------------------------------------------------------------------------
def _norm(s: str) -> str:
    s = "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c))
    return s.strip().upper()


def find_col(df: pd.DataFrame, *candidates: str):
    """Encuentra una columna por nombre, tolerante a acentos/mayusculas."""
    norm_map = {_norm(c): c for c in df.columns}
    for cand in candidates:
        if _norm(cand) in norm_map:
            return norm_map[_norm(cand)]
    return None


def parse_date(x):
    if pd.isna(x) or str(x).strip() in ("-", "", "nan"):
        return pd.NaT
    return pd.to_datetime(str(x).strip(), dayfirst=True, errors="coerce")


def clean_consultor(x) -> str:
    if pd.isna(x):
        return ""
    s = re.sub(r",\s*$", "", str(x)).strip()
    if s in ("-", "", "nan"):
        return ""
    if "," in s:                      # si vinieran varios, tomamos el primero (Consultor 1)
        s = s.split(",")[0].strip()
    return s


def parse_empleados(x):
    if pd.isna(x) or str(x).strip() in ("-", "", "nan"):
        return None
    try:
        return int(float(str(x).strip().replace(".", "").replace(",", "")))
    except ValueError:
        return None


# --------------------------------------------------------------------------
# 3. Carga de proyectos (hoy: Excel | futuro: API Zoho)
# --------------------------------------------------------------------------
def cargar_proyectos(uploaded) -> pd.DataFrame:
    """
    Lee la hoja del export que contenga los proyectos. Para enchufar Zoho
    a futuro, reemplazar esta funcion por la llamada a la API que devuelva
    un DataFrame con las mismas columnas normalizadas mas abajo.
    """
    sheets = pd.read_excel(uploaded, sheet_name=None, dtype=str)
    best, best_score = None, -1
    for name, df in sheets.items():
        df.columns = [str(c).strip() for c in df.columns]
        score = sum(find_col(df, c) is not None for c in
                    ["ID DEL PROYECTO", "FECHA FINAL", "CANTIDAD DE EMPLEADOS", "CONSULTOR 1"])
        if score > best_score:
            best, best_score = df, score
    return best


# --------------------------------------------------------------------------
# 4. Configuracion (sidebar)
# --------------------------------------------------------------------------
st.title("Cálculo de Bono Consultores")
st.caption("Página aislada · v1 sobre export de Zoho · factor de fechas configurable")

with st.sidebar:
    st.header("⚙️ Parámetros")

    st.subheader("Período (sobre Fecha Final)")
    f_ini = st.date_input("Desde", value=date(2026, 1, 1), format="DD/MM/YYYY")
    f_fin = st.date_input("Hasta", value=date(2026, 3, 31), format="DD/MM/YYYY")
    solo_cerrados = st.checkbox("Solo proyectos Cerrados", value=True)

    st.subheader("Puntos por dotación")
    rangos = [(1, 100), (101, 250), (251, 500), (501, 1000), (1001, 3000), (3001, 10**9)]
    labels = ["1-100", "101-250", "251-500", "501-1000", "1001-3000", "3001+"]
    defaults_pts = [0.75, 1.15, 1.75, 2.20, 2.70, 5.00]
    puntos_rango = {}
    for lbl, dv in zip(labels, defaults_pts):
        puntos_rango[lbl] = st.number_input(f"Puntos {lbl}", value=float(dv), step=0.05, format="%.2f")

    with st.expander("Puntos por tipo de servicio (sesiones)"):
        SESIONES_DEFAULT = {
            "Portal": 0.25,
            "Centralización desde cero": 0.95,
            "Centralización M+": 0.45,
            "Migración Archivos": 1.00,
            "Organigrama": 0.85,
            "Escala y grados": 1.00,
            "Repaso": 0.55,
            "CDR": 0.65,
            "Tratos": 0.75,
            "LME": 0.35,
            "Geo victoria": 0.15,
            "Conexion a la DT": 0.10,
            "Comunicaciones y Bene": 0.05,
            "Trabajo Adm": 0.075,
            "Gecos": 1.25,
            "Smart Rex": 0.20,
        }
        puntos_sesion = {}
        for nombre, dv in SESIONES_DEFAULT.items():
            puntos_sesion[nombre] = st.number_input(
                f"Puntos {nombre}", value=float(dv), step=0.05, format="%.3f", key=f"ses_{nombre}"
            )
    # lookup tolerante a acentos/mayúsculas
    puntos_sesion_norm = {_norm(k): v for k, v in puntos_sesion.items()}

    st.subheader("Tramos de fechas (% por días de atraso)")
    c1 = st.number_input("Corte 1 (días)", value=3, step=1)
    c2 = st.number_input("Corte 2 (días)", value=7, step=1)
    p0 = st.number_input("% a tiempo o antes (≤0)", value=100, step=5) / 100
    p1 = st.number_input(f"% 1 a {c1} días", value=75, step=5) / 100
    p2 = st.number_input(f"% {c1+1} a {c2} días", value=50, step=5) / 100
    p3 = st.number_input(f"% más de {c2} días", value=25, step=5) / 100

    st.subheader("Factor de fechas")
    peso_a = st.slider("Peso Métrica A (cierre vs facturación)", 0.0, 1.0, 0.40, 0.05)
    peso_b = round(1 - peso_a, 2)
    st.caption(f"Peso Métrica B (cierre vs plan Fin 1) = {peso_b:.2f}")
    piso = st.slider("Piso del factor", 0.0, 1.0, 0.70, 0.05)

    st.subheader("Tramos de bono (puntos → monto)")
    u_alto = st.number_input("Umbral alto (puntos)", value=15.0, step=0.5)
    u_medio = st.number_input("Umbral medio (puntos)", value=10.0, step=0.5)
    u_bajo = st.number_input("Umbral bajo (puntos)", value=5.0, step=0.5)
    m_alto = st.number_input("Monto ≥ alto", value=200000, step=10000)
    m_medio = st.number_input("Monto medio", value=150000, step=10000)
    m_bajo = st.number_input("Monto bajo", value=100000, step=10000)
    m_min = st.number_input("Monto mínimo (≤ bajo)", value=50000, step=10000)


def pct_por_delta(delta):
    if pd.isna(delta):
        return None
    if delta <= 0:
        return p0
    if delta <= c1:
        return p1
    if delta <= c2:
        return p2
    return p3


def bracket_label(emp):
    if emp is None:
        return None
    for (lo, hi), lbl in zip(rangos, labels):
        if lo <= emp <= hi:
            return lbl
    return None


def monto_por_puntos(pts):
    if pts >= u_alto:
        return m_alto, "Alto"
    if pts >= u_medio:
        return m_medio, "Medio"
    if pts > u_bajo:
        return m_bajo, "Bajo"
    return m_min, "Mínimo"


def puntos_por_tipo(tipo):
    if tipo is None or str(tipo).strip() in ("", "-", "nan"):
        return 0.0
    return puntos_sesion_norm.get(_norm(tipo), 0.0)


# --------------------------------------------------------------------------
# 5. Carga de datos
# --------------------------------------------------------------------------
uploaded = st.file_uploader("Sube el export de proyectos de Zoho (.xlsx)", type=["xlsx"])
adic_text = st.text_area(
    "Ajuste manual de puntos por consultor (opcional) — formato `Nombre: puntos`, uno por línea. "
    "Úsalo solo para correcciones puntuales; las sesiones se calculan solas desde Tipo de servicio.",
    placeholder="Sebastian Leon: 1.5",
    height=70,
)

adicionales = {}
for line in adic_text.splitlines():
    if ":" in line:
        n, v = line.split(":", 1)
        try:
            adicionales[clean_consultor(n)] = float(v.strip().replace(",", "."))
        except ValueError:
            pass

if uploaded is None:
    st.info("Esperando el archivo para calcular.")
    st.stop()

raw = cargar_proyectos(uploaded)

col_id = find_col(raw, "ID DEL PROYECTO")
col_emp = find_col(raw, "CANTIDAD DE EMPLEADOS")
col_cons = find_col(raw, "CONSULTOR 1")
col_estado = find_col(raw, "ESTADO")
col_final = find_col(raw, "FECHA FINAL")
col_fact = find_col(raw, "FECHA FACTURACION", "FECHA FACTURACIÓN")
col_fin1 = find_col(raw, "FECHA FIN 1")
col_grupo = find_col(raw, "GRUPO DE PROYECTOS")
col_tipo = find_col(raw, "TIPO DE SERVICIO")

faltan = [n for n, c in {
    "ID": col_id, "Empleados": col_emp, "Consultor 1": col_cons,
    "Fecha Final": col_final}.items() if c is None]
if faltan:
    st.error(f"No encontré columnas requeridas: {', '.join(faltan)}")
    st.stop()

df = pd.DataFrame()
df["id"] = raw[col_id]
df["consultor"] = raw[col_cons].apply(clean_consultor)
df["empleados"] = raw[col_emp].apply(parse_empleados)
df["estado"] = raw[col_estado] if col_estado else ""
df["f_final"] = raw[col_final].apply(parse_date)
df["f_fact"] = raw[col_fact].apply(parse_date) if col_fact else pd.NaT
df["f_fin1"] = raw[col_fin1].apply(parse_date) if col_fin1 else pd.NaT
df["grupo"] = raw[col_grupo].astype(str).str.strip() if col_grupo else ""
df["tipo"] = raw[col_tipo].astype(str).str.strip() if col_tipo else ""

# --- Selección de grupos: dotación (por empleados) y sesiones (por tipo) ---
grupos_disponibles = sorted([g for g in df["grupo"].unique() if g and g != "nan"])
PRESETS = {
    "Remuneraciones / DO": {
        "dotacion": ["REX-PROYECTO VENDEDOR REXMAS", "REX-PROYECTO VENDEDOR MANAGER", "REX-DO"],
        "sesiones": ["REX-SERVICIO ESPECIAL VENDEDOR REXMAS", "REX-SERVICIO ESPECIAL VENDEDOR MANAGER",
                     "REX-ACADEMIA"],
    },
    "Asistencia": {
        "dotacion": ["BNS-ASISTENCIA/CASINO"],
        "sesiones": ["BNS-SERVICIO ESPECIAL REXMAS", "BNS-SERVICIO ESPECIAL MANAGER", "BNS-ADICIONALES"],
    },
    "Personalizado": {"dotacion": [], "sesiones": []},
}
tipo_bono = st.radio("Bono a calcular", list(PRESETS.keys()), horizontal=True)
def_dot = [g for g in PRESETS[tipo_bono]["dotacion"] if g in grupos_disponibles]
def_ses = [g for g in PRESETS[tipo_bono]["sesiones"] if g in grupos_disponibles]

grupos_dot = st.multiselect(
    "Grupos que puntúan por DOTACIÓN (cantidad de empleados)",
    options=grupos_disponibles, default=def_dot,
)
grupos_ses = st.multiselect(
    "Grupos que puntúan por SESIÓN (campo Tipo de servicio)",
    options=grupos_disponibles, default=def_ses,
    help="Requiere el campo 'Tipo de servicio' cargado en Zoho. Sin ese campo, estos suman 0.",
)
if not grupos_dot and not grupos_ses:
    st.warning("Selecciona al menos un grupo (dotación o sesión).")
    st.stop()
if grupos_ses and col_tipo is None:
    st.info("Aún no encuentro la columna 'Tipo de servicio' en el export; las sesiones sumarán 0 hasta que la cargues en Zoho.")

# filtros base (período + estado + consultor)
base = df["f_final"].between(pd.Timestamp(f_ini), pd.Timestamp(f_fin))
if solo_cerrados and col_estado:
    base &= df["estado"].apply(lambda s: _norm(s) == "CERRADO")
base &= df["consultor"].astype(bool)
df = df[base & df["grupo"].isin(grupos_dot + grupos_ses)].copy()

if df.empty:
    st.warning("No hay proyectos en el período/filtros elegidos.")
    st.stop()

# clasificación y puntos por fila
df["es_dotacion"] = df["grupo"].isin(grupos_dot)
df["rango"] = df.apply(lambda r: bracket_label(r["empleados"]) if r["es_dotacion"] else None, axis=1)
df["puntos_proy"] = df.apply(
    lambda r: puntos_rango.get(r["rango"], 0.0) if r["es_dotacion"] else 0.0, axis=1)
df["puntos_ses"] = df.apply(
    lambda r: 0.0 if r["es_dotacion"] else puntos_por_tipo(r["tipo"]), axis=1)

# las métricas de fecha solo aplican a proyectos de dotación (plazos comprometidos)
df["delta_A"] = (df["f_final"] - df["f_fact"]).dt.days.where(df["es_dotacion"])
df["delta_B"] = (df["f_final"] - df["f_fin1"]).dt.days.where(df["es_dotacion"])
df["pct_A"] = df["delta_A"].apply(pct_por_delta)
df["pct_B"] = df["delta_B"].apply(pct_por_delta)


# --------------------------------------------------------------------------
# 6. Agregado por consultor
# --------------------------------------------------------------------------
def factor_fechas(g):
    a = g["pct_A"].dropna()
    b = g["pct_B"].dropna()
    ma = a.mean() if len(a) else None
    mb = b.mean() if len(b) else None
    if ma is None and mb is None:
        return 1.0, ma, mb            # sin datos de fecha: no se penaliza
    if ma is None:
        f = mb
    elif mb is None:
        f = ma
    else:
        f = peso_a * ma + peso_b * mb
    return max(piso, f), ma, mb


filas = []
for cons, g in df.groupby("consultor"):
    pts_proy = g["puntos_proy"].sum()
    pts_ses = g["puntos_ses"].sum()
    pts_adic = adicionales.get(cons, 0.0)        # ajuste manual opcional
    total_pts = pts_proy + pts_ses + pts_adic
    monto_base, tier = monto_por_puntos(total_pts)
    factor, ma, mb = factor_fechas(g)
    n_dot = int(g["es_dotacion"].sum())
    filas.append({
        "Consultor": cons,
        "Proy. dotación": n_dot,
        "Sesiones": int(len(g) - n_dot),
        "Puntos proyectos": round(pts_proy, 2),
        "Puntos sesiones": round(pts_ses, 2),
        "Ajuste manual": round(pts_adic, 2),
        "Total puntos": round(total_pts, 2),
        "Tramo": tier,
        "Monto base": int(monto_base),
        "% Métrica A": round(ma * 100, 1) if ma is not None else None,
        "% Métrica B": round(mb * 100, 1) if mb is not None else None,
        "Factor fechas": round(factor, 3),
        "Bono final": int(round(monto_base * factor)),
    })

resumen = pd.DataFrame(filas).sort_values("Bono final", ascending=False).reset_index(drop=True)


# --------------------------------------------------------------------------
# 7. Salida
# --------------------------------------------------------------------------
c1m, c2m, c3m = st.columns(3)
c1m.metric("Consultores", len(resumen))
c2m.metric("Proyectos considerados", len(df))
c3m.metric("Total a pagar", f"${resumen['Bono final'].sum():,.0f}".replace(",", "."))

n_sin_fact = int(df["delta_A"].isna().sum())
n_sin_fin1 = int(df["delta_B"].isna().sum())
if n_sin_fact or n_sin_fin1:
    st.warning(
        f"Datos incompletos en Zoho: {n_sin_fact} proyectos sin fecha de facturación "
        f"y {n_sin_fin1} sin Fin 1 (esos se excluyen de su métrica, no penalizan). "
        "Para automatizar de forma confiable, vuelve obligatorios esos campos al cerrar."
    )

st.subheader("Resumen por consultor")
st.dataframe(resumen, use_container_width=True, hide_index=True)

st.subheader("Detalle por proyecto")
detalle = df[["id", "consultor", "grupo", "tipo", "empleados", "rango",
              "puntos_proy", "puntos_ses", "f_fact", "f_final", "f_fin1",
              "delta_A", "delta_B", "pct_A", "pct_B"]].copy()
for c in ["f_fact", "f_final", "f_fin1"]:
    detalle[c] = detalle[c].dt.strftime("%d/%m/%Y")
detalle["pct_A"] = (detalle["pct_A"] * 100).round(0)
detalle["pct_B"] = (detalle["pct_B"] * 100).round(0)
st.dataframe(detalle, use_container_width=True, hide_index=True)

# export
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as xw:
    resumen.to_excel(xw, sheet_name="Resumen", index=False)
    detalle.to_excel(xw, sheet_name="Detalle", index=False)
st.download_button(
    "⬇️ Descargar Excel del cálculo",
    data=buf.getvalue(),
    file_name=f"bono_consultores_{f_ini:%Y%m%d}_{f_fin:%Y%m%d}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ------------------------------------------------------------------
# NOTA para ocultar del menu lateral:
# Streamlit muestra automaticamente lo que este en pages/. Para que esta
# pagina no aparezca en el nav, una opcion es manejar la navegacion con
# st.navigation() en tu archivo principal y agregar esta pagina solo de
# forma condicional. Como ya tiene candado, el contenido sensible queda
# protegido aunque el nombre sea visible.
# ------------------------------------------------------------------
