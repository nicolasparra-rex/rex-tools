import streamlit as st
import requests
import pandas as pd
import json
from datetime import datetime, date
from calendar import monthrange

st.set_page_config(page_title="Dashboard", page_icon="📊", layout="wide")

# ── AUTH ──────────────────────────────────────────────────────────────────────

@st.cache_data(ttl=3000, show_spinner=False)
def get_access_token(refresh_token, client_id, client_secret):
    r = requests.post("https://accounts.zoho.com/oauth/v2/token", params={
        "refresh_token": refresh_token,
        "client_id":     client_id,
        "client_secret": client_secret,
        "grant_type":    "refresh_token",
    })
    data = r.json()
    return data.get("access_token")


@st.cache_data(ttl=600, show_spinner=False)
def get_all_projects(access_token, portal_id):
    url = f"https://projectsapi.zoho.com/restapi/portal/{portal_id}/projects/"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    all_projects = []
    index = 1
    while True:
        r = requests.get(url, headers=headers, params={"range": 100, "index": index})
        batch = r.json().get("projects", [])
        if not batch:
            break
        all_projects.extend(batch)
        if len(batch) < 100:
            break
        index += 100
    return all_projects

# ── HELPERS ───────────────────────────────────────────────────────────────────

def parse_custom_fields(custom_fields):
    result = {}
    if not isinstance(custom_fields, list):
        return result
    for item in custom_fields:
        if isinstance(item, dict):
            for k, v in item.items():
                result[k] = v
    return result

def cf(fields, *keys):
    for k in keys:
        if k in fields and fields[k] not in (None, "", "false", False):
            val = fields[k]
            if isinstance(val, str) and val.startswith("["):
                try:
                    parsed = json.loads(val)
                    return ", ".join(parsed) if isinstance(parsed, list) else val
                except Exception:
                    pass
            return val
    return ""

def parse_date(d):
    if not d:
        return None
    for fmt in ("%m-%d-%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(d, fmt).date()
        except Exception:
            pass
    return None

def safe_int(v):
    try:
        return int(str(v).replace(",", "").strip())
    except Exception:
        return 0

def build_df(projects):
    rows = []
    for p in projects:
        cfields     = parse_custom_fields(p.get("custom_fields", []))
        status_name = p.get("custom_status_name") or p.get("status", "")
        empleados   = safe_int(cf(cfields, "Cantidad de empleados"))
        end_date    = parse_date(p.get("end_date", ""))
        ffact       = parse_date(cf(cfields, "Fecha Facturación", "Fecha Facturacion"))
        rows.append({
            "nombre":      p.get("name", ""),
            "key":         p.get("key", ""),
            "status":      status_name,
            "grupo":       p.get("group_name", ""),
            "jefe":        p.get("owner_name", ""),
            "consultor":   cf(cfields, "Consultor 1"),
            "empleados":   empleados,
            "end_date":    end_date,
            "ffact":       ffact,
            "created_date": parse_date(p.get("created_date", "")),
            "plan":        cf(cfields, "Plan Contratado"),
            "razon":       cf(cfields, "Razón social"),
        })
    return pd.DataFrame(rows)

def fmt_num(n):
    return f"{int(n):,}".replace(",", ".")

# ── UI ────────────────────────────────────────────────────────────────────────

st.title("📊 Dashboard Rex+")
st.caption("Datos en tiempo real desde Zoho Projects")

portal_id = st.secrets.get("ZOHO_PORTAL_ID", "757079135")

_, col_btn = st.columns([6, 1])
with col_btn:
    if st.button("🔄 Actualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

with st.spinner("Conectando con Zoho..."):
    token = get_access_token(
        st.secrets["ZOHO_REFRESH_TOKEN"],
        st.secrets["ZOHO_CLIENT_ID"],
        st.secrets["ZOHO_CLIENT_SECRET"],
    )

if not token:
    st.error("❌ No se pudo obtener el token.")
    st.stop()

with st.spinner("Cargando proyectos..."):
    projects = get_all_projects(token, portal_id)

if not projects:
    st.warning("No se encontraron proyectos.")
    st.stop()

df = build_df(projects)

# ── FILTROS SUPERIORES ────────────────────────────────────────────────────────
st.divider()
ff1, ff2, ff3 = st.columns(3)
with ff1:
    grupos_opts = ["Todos"] + sorted([x for x in df["grupo"].dropna().unique() if x])
    filtro_grupo = st.selectbox("Grupo proyecto", grupos_opts)
with ff2:
    jefes_opts = ["Todos"] + sorted([x for x in df["jefe"].dropna().unique() if x])
    filtro_jefe = st.selectbox("Jefe de proyecto", jefes_opts)
with ff3:
    consultores_opts = ["Todos"] + sorted([x for x in df["consultor"].dropna().unique() if x])
    filtro_consultor = st.selectbox("Consultor", consultores_opts)

# Aplicar filtros
dff = df.copy()
if filtro_grupo != "Todos":
    dff = dff[dff["grupo"] == filtro_grupo]
if filtro_jefe != "Todos":
    dff = dff[dff["jefe"] == filtro_jefe]
if filtro_consultor != "Todos":
    dff = dff[dff["consultor"] == filtro_consultor]

# ── CÁLCULOS ──────────────────────────────────────────────────────────────────
hoy        = date.today()
mes_actual = hoy.month
anio_actual= hoy.year

def mes_rango(offset=0):
    m = mes_actual + offset
    a = anio_actual
    while m > 12:
        m -= 12; a += 1
    while m < 1:
        m += 12; a -= 1
    ultimo = monthrange(a, m)[1]
    return date(a, m, 1), date(a, m, ultimo)

mes_ant_ini,  mes_ant_fin  = mes_rango(-1)
mes_act_ini,  mes_act_fin  = mes_rango(0)
mes_sig_ini,  mes_sig_fin  = mes_rango(1)
_, tres_meses_ini = mes_rango(-2)  # fin del mes de hace 2 meses hacia atrás
tres_meses_ini = mes_rango(-2)[0]  # inicio del mes de hace 2 meses

# Grupos Rex+ para KPIs de cartera
GRUPOS_REX = ["rex-proyectos vendedor rexmas", "rex-proyectos vendedor manager"]
dff_rex = dff[dff["grupo"].str.lower().isin([g.lower() for g in GRUPOS_REX])]

# Status
STATUS_DETENIDO = ["detenido por cliente", "detenido por comercial"]
def es_detenido(s): return any(d in s.lower() for d in STATUS_DETENIDO)
def es_activo(s):   return not es_detenido(s)

activos   = dff_rex[dff_rex["status"].apply(es_activo)]
detenidos = dff_rex[dff_rex["status"].apply(es_detenido)]
total_rex = dff_rex

emp_activos   = activos["empleados"].sum()
emp_detenidos = detenidos["empleados"].sum()
emp_total     = total_rex["empleados"].sum()

# Últimos 3 meses = proyectos creados en los últimos 3 meses
ult3     = dff_rex[dff_rex["created_date"].apply(lambda d: d is not None and tres_meses_ini <= d <= hoy)]
emp_ult3 = ult3["empleados"].sum()

# Salidas planificadas por end_date (sobre dff filtrado por usuario, no solo rex)
def en_mes(d, ini, fin): return d is not None and ini <= d <= fin

sal_ant    = dff[dff["end_date"].apply(lambda d: en_mes(d, mes_ant_ini, mes_ant_fin))]
sal_act    = dff[dff["end_date"].apply(lambda d: en_mes(d, mes_act_ini, mes_act_fin))]
sal_sig    = dff[dff["end_date"].apply(lambda d: en_mes(d, mes_sig_ini, mes_sig_fin))]
sal_sin_ag = dff[dff["end_date"].isna() | dff["end_date"].isnull()]

emp_sal_ant    = sal_ant["empleados"].sum()
emp_sal_act    = sal_act["empleados"].sum()
emp_sal_sig    = sal_sig["empleados"].sum()
emp_sal_sin_ag = sal_sin_ag["empleados"].sum()

# ── KPIs CLIENTES ENTRADA ─────────────────────────────────────────────────────
st.markdown("## 🏢 Clientes Entrada / Salida")
st.divider()

col_izq, col_der = st.columns([1, 1])

with col_izq:
    st.markdown("### 📥 Cartera actual")
    r1c1, r1c2 = st.columns(2)
    r1c1.metric("Activos",   fmt_num(len(activos)))
    r1c2.metric("Empleados", fmt_num(emp_activos))

    r2c1, r2c2 = st.columns(2)
    r2c1.metric("Detenidos",  fmt_num(len(detenidos)))
    r2c2.metric("Empleados",  fmt_num(emp_detenidos))

    r3c1, r3c2 = st.columns(2)
    r3c1.metric("Total",     fmt_num(len(total_df)))
    r3c2.metric("Empleados", fmt_num(emp_total))

    st.markdown("")
    r4c1, r4c2 = st.columns(2)
    r4c1.metric("Últimos 3 meses", fmt_num(len(ult3)))
    r4c2.metric("Empleados",       fmt_num(emp_ult3))

with col_der:
    st.markdown("### 📤 Salidas planificadas (end date)")
    s1c1, s1c2 = st.columns(2)
    s1c1.metric("Mes anterior",  fmt_num(len(sal_ant)))
    s1c2.metric("Empleados",     fmt_num(emp_sal_ant))

    s2c1, s2c2 = st.columns(2)
    s2c1.metric("Mes actual",    fmt_num(len(sal_act)))
    s2c2.metric("Empleados",     fmt_num(emp_sal_act))

    s3c1, s3c2 = st.columns(2)
    s3c1.metric("Mes siguiente", fmt_num(len(sal_sig)))
    s3c2.metric("Empleados",     fmt_num(emp_sal_sig))

    s4c1, s4c2 = st.columns(2)
    s4c1.metric("Sin fecha",     fmt_num(len(sal_sin_ag)))
    s4c2.metric("Empleados",     fmt_num(emp_sal_sin_ag))

# ── KPI: PROYECTOS POR CONSULTOR ──────────────────────────────────────────────
st.divider()
st.markdown("### 👤 Proyectos por Consultor")

if dff["consultor"].any():
    cons_df = (
        dff[dff["consultor"] != ""]
        .groupby("consultor")
        .agg(
            Proyectos  = ("nombre", "count"),
            Empleados  = ("empleados", "sum"),
            Activos    = ("status", lambda x: sum(es_activo(s) for s in x)),
            Detenidos  = ("status", lambda x: sum(es_detenido(s) for s in x)),
        )
        .reset_index()
        .sort_values("Proyectos", ascending=False)
        .rename(columns={"consultor": "Consultor"})
    )
    cons_df["Empleados"] = cons_df["Empleados"].apply(fmt_num)
    st.dataframe(cons_df, use_container_width=True, hide_index=True)
else:
    st.info("No hay datos de consultor disponibles.")

# ── KPI: PROYECTOS POR GRUPO ──────────────────────────────────────────────────
st.divider()
st.markdown("### 🗂️ Proyectos por Grupo")

grupo_df = (
    dff[dff["grupo"] != ""]
    .groupby("grupo")
    .agg(
        Proyectos = ("nombre", "count"),
        Empleados = ("empleados", "sum"),
        Activos   = ("status", lambda x: sum(es_activo(s) for s in x)),
        Detenidos = ("status", lambda x: sum(es_detenido(s) for s in x)),
    )
    .reset_index()
    .sort_values("Proyectos", ascending=False)
    .rename(columns={"grupo": "Grupo"})
)
grupo_df["Empleados"] = grupo_df["Empleados"].apply(fmt_num)
st.dataframe(grupo_df, use_container_width=True, hide_index=True)

# ── KPI: PROYECTOS POR PLAN ───────────────────────────────────────────────────
st.divider()
st.markdown("### 📋 Proyectos por Plan")

plan_df = (
    dff[dff["plan"] != ""]
    .groupby("plan")
    .agg(
        Proyectos = ("nombre", "count"),
        Empleados = ("empleados", "sum"),
    )
    .reset_index()
    .sort_values("Proyectos", ascending=False)
    .rename(columns={"plan": "Plan"})
)
plan_df["Empleados"] = plan_df["Empleados"].apply(fmt_num)
st.dataframe(plan_df, use_container_width=True, hide_index=True)

# ── SALIDAS DETALLE ───────────────────────────────────────────────────────────
st.divider()
st.markdown("### 📤 Detalle salidas planificadas")

tab1, tab2, tab3 = st.tabs([
    f"Mes anterior ({mes_ant_ini.strftime('%B %Y')})",
    f"Mes actual ({mes_act_ini.strftime('%B %Y')})",
    f"Mes siguiente ({mes_sig_ini.strftime('%B %Y')})",
])

def tabla_salidas(df_sal):
    if df_sal.empty:
        st.info("Sin proyectos en este período.")
        return
    cols = ["key", "nombre", "razon", "consultor", "empleados", "end_date"]
    d = df_sal[cols].copy()
    d["end_date"] = d["end_date"].apply(lambda x: x.strftime("%d/%m/%Y") if x else "–")
    d.columns = ["Clave", "Proyecto", "Razón Social", "Consultor", "Empleados", "Fecha Término"]
    st.dataframe(d, use_container_width=True, hide_index=True)

with tab1: tabla_salidas(sal_ant)
with tab2: tabla_salidas(sal_act)
with tab3: tabla_salidas(sal_sig)
