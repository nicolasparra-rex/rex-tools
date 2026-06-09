import streamlit as st
import requests
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Zoho Proyectos", page_icon="📋", layout="wide")

# ── AUTENTICACIÓN ─────────────────────────────────────────────────────────────

@st.cache_data(ttl=3000, show_spinner=False)
def get_access_token(refresh_token, client_id, client_secret):
    """Obtiene un Access Token. Cacheado 50 min (el token expira en 60)."""
    url = "https://accounts.zoho.com/oauth/v2/token"
    params = {
        "refresh_token": refresh_token,
        "client_id":     client_id,
        "client_secret": client_secret,
        "grant_type":    "refresh_token",
    }
    r = requests.post(url, params=params)
    data = r.json()
    if "access_token" in data:
        return data["access_token"]
    return None


@st.cache_data(ttl=600, show_spinner=False)
def get_projects(access_token, portal_id):
    """Obtiene todos los proyectos activos. Cacheado 10 minutos."""
    url = f"https://projectsapi.zoho.com/restapi/portal/{portal_id}/projects/"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    params = {"status": "active", "range": 100}
    r = requests.get(url, headers=headers, params=params)
    data = r.json()
    return data.get("projects", [])


@st.cache_data(ttl=600, show_spinner=False)
def get_tasks(access_token, portal_id, project_id):
    """Obtiene las tareas de un proyecto. Cacheado 10 minutos."""
    url = f"https://projectsapi.zoho.com/restapi/portal/{portal_id}/projects/{project_id}/tasks/"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    params = {"range": 100}
    r = requests.get(url, headers=headers, params=params)
    data = r.json()
    return data.get("tasks", [])


# ── HELPERS ───────────────────────────────────────────────────────────────────

STATUS_COLORS = {
    "inicio sin agenda": "🟠",
    "en curso":          "🔵",
    "reunion ko":        "🟣",
    "completado":        "🟢",
    "cerrado":           "⚫",
}

def status_icon(name):
    key = (name or "").lower()
    for k, v in STATUS_COLORS.items():
        if k in key:
            return v
    return "⚪"

def get_status_name(p):
    s = p.get("status", "")
    if isinstance(s, dict):
        return s.get("name", "–")
    if isinstance(s, str) and s:
        return s
    return "–"

def fmt_date(d):
    if not d:
        return "–"
    try:
        return datetime.strptime(d, "%m-%d-%Y").strftime("%d/%m/%Y")
    except Exception:
        return d

def build_df(projects):
    rows = []
    for p in projects:
        status_name = get_status_name(p)
        task_count  = p.get("task_count", {}) or {}
        rows.append({
            "Clave":        p.get("id_string", ""),
            "Proyecto":     p.get("name", ""),
            "Estado":       status_icon(status_name) + " " + status_name,
            "Consultor":    p.get("owner_name", "–"),
            "Plan":         p.get("plan_contratado", "–"),
            "Empleados":    p.get("cantidad_de_empleados", "–"),
            "Vendedor":     p.get("vendedor", "–"),
            "Razón Social": p.get("razon_social", ""),
            "Tareas ✅":    task_count.get("closed", 0),
            "Tareas 🔓":    task_count.get("open", 0),
            "Creado":       fmt_date(p.get("created_date", "")),
            "_id":          str(p.get("id", "")),
            "_nombre_raw":  p.get("name", ""),
            "_status_raw":  status_name,
            "_plan_raw":    p.get("plan_contratado", ""),
            "_owner_raw":   p.get("owner_name", ""),
        })
    return pd.DataFrame(rows)


# ── UI ────────────────────────────────────────────────────────────────────────

st.title("📋 Zoho Proyectos")
st.caption("Proyectos activos del portal Rex+ · datos en tiempo real")

portal_id = st.secrets.get("ZOHO_PORTAL_ID", "757079135")

col_title, col_btn = st.columns([6, 1])
with col_btn:
    if st.button("🔄 Actualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# Obtener token
with st.spinner("Conectando con Zoho Projects..."):
    token = get_access_token(
        st.secrets["ZOHO_REFRESH_TOKEN"],
        st.secrets["ZOHO_CLIENT_ID"],
        st.secrets["ZOHO_CLIENT_SECRET"],
    )

if not token:
    st.error("❌ No se pudo obtener el token. Revisa los secrets en Streamlit.")
    st.stop()

# Obtener proyectos
with st.spinner("Cargando proyectos..."):
    projects = get_projects(token, portal_id)

if not projects:
    st.warning("No se encontraron proyectos activos.")
    st.stop()

df = build_df(projects)

# ── KPIs ──────────────────────────────────────────────────────────────────────
st.divider()
total      = len(df)
sin_agenda = df["_status_raw"].str.lower().str.contains("inicio sin agenda", na=False).sum()
en_curso   = df["_status_raw"].str.lower().str.contains("en curso", na=False).sum()
otros      = total - sin_agenda - en_curso
t_abiertas = int(df["Tareas 🔓"].sum())

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total proyectos", total)
k2.metric("Sin agendar 🟠", int(sin_agenda))
k3.metric("En curso 🔵", int(en_curso))
k4.metric("Otras etapas", int(otros))
k5.metric("Tareas abiertas", t_abiertas)

st.divider()

# ── FILTROS ───────────────────────────────────────────────────────────────────
st.subheader("🔍 Filtros")
fc1, fc2, fc3, fc4 = st.columns(4)

with fc1:
    search = st.text_input("Buscar cliente o proyecto", placeholder="Nombre, razón social...")
with fc2:
    estados = ["Todos"] + sorted(df["_status_raw"].dropna().unique().tolist())
    filtro_estado = st.selectbox("Estado", estados)
with fc3:
    planes = ["Todos"] + sorted([x for x in df["_plan_raw"].dropna().unique() if x])
    filtro_plan = st.selectbox("Plan", planes)
with fc4:
    consultores = ["Todos"] + sorted(df["_owner_raw"].dropna().unique().tolist())
    filtro_consultor = st.selectbox("Consultor", consultores)

mask = pd.Series([True] * len(df))
if search:
    mask &= (
        df["_nombre_raw"].str.contains(search, case=False, na=False) |
        df["Razón Social"].str.contains(search, case=False, na=False)
    )
if filtro_estado != "Todos":
    mask &= df["_status_raw"] == filtro_estado
if filtro_plan != "Todos":
    mask &= df["_plan_raw"] == filtro_plan
if filtro_consultor != "Todos":
    mask &= df["_owner_raw"] == filtro_consultor

df_filtered = df[mask].reset_index(drop=True)
st.caption(f"{len(df_filtered)} proyectos encontrados")

# ── TABLA ─────────────────────────────────────────────────────────────────────
st.divider()
st.subheader("📊 Listado de proyectos")

cols_show = ["Clave", "Proyecto", "Estado", "Consultor", "Plan", "Vendedor", "Empleados", "Tareas ✅", "Tareas 🔓", "Creado"]
st.dataframe(
    df_filtered[cols_show],
    use_container_width=True,
    hide_index=True,
    column_config={
        "Clave":     st.column_config.TextColumn(width="small"),
        "Proyecto":  st.column_config.TextColumn(width="large"),
        "Estado":    st.column_config.TextColumn(width="medium"),
        "Tareas ✅": st.column_config.NumberColumn(width="small"),
        "Tareas 🔓": st.column_config.NumberColumn(width="small"),
    }
)

# ── DETALLE DE PROYECTO ───────────────────────────────────────────────────────
st.divider()
st.subheader("📌 Detalle de proyecto")

project_names = ["Selecciona un proyecto..."] + df_filtered["Proyecto"].tolist()
selected = st.selectbox("Proyecto", project_names, label_visibility="collapsed")

if selected != "Selecciona un proyecto...":
    row = df_filtered[df_filtered["Proyecto"] == selected].iloc[0]
    project_id = row["_id"]
    proj_data  = next((p for p in projects if str(p.get("id")) == project_id), {})

    d1, d2, d3 = st.columns(3)
    with d1:
        st.markdown("**Razón social**")
        st.write(proj_data.get("razon_social") or "–")
        st.markdown("**RUT**")
        st.write(proj_data.get("rut_empresa") or "–")
        st.markdown("**Contacto**")
        st.write(proj_data.get("nombre_del_contacto") or "–")
    with d2:
        st.markdown("**Plan contratado**")
        st.write(proj_data.get("plan_contratado") or "–")
        st.markdown("**Módulos vendidos**")
        modulos = proj_data.get("modulo_vendido", [])
        st.write(", ".join(modulos) if isinstance(modulos, list) else str(modulos or "–"))
        st.markdown("**Vendedor**")
        st.write(proj_data.get("vendedor") or "–")
    with d3:
        st.markdown("**Empleados**")
        st.write(proj_data.get("cantidad_de_empleados") or "–")
        st.markdown("**Empresas**")
        st.write(proj_data.get("cantidad_de_empresas") or "–")
        st.markdown("**Fecha venta**")
        st.write(proj_data.get("fecha_de_venta") or "–")

    # Tareas
    st.markdown("---")
    st.markdown("**🗂️ Tareas del proyecto**")
    with st.spinner("Cargando tareas..."):
        tasks = get_tasks(token, portal_id, project_id)

    if tasks:
        task_rows = []
        for t in tasks:
            s = t.get("status", {})
            status_t = s.get("name", "–") if isinstance(s, dict) else str(s or "–")
            owners = t.get("details", {}).get("owners", []) if t.get("details") else []
            responsable = owners[0].get("name", "–") if owners else "–"
            task_rows.append({
                "Tarea":        t.get("name", ""),
                "Estado":       status_t,
                "% avance":     int(t.get("percent_complete", 0) or 0),
                "Responsable":  responsable,
                "Vencimiento":  fmt_date(t.get("end_date", "")),
            })
        df_tasks = pd.DataFrame(task_rows)
        st.dataframe(
            df_tasks,
            use_container_width=True,
            hide_index=True,
            column_config={
                "% avance": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%d%%"),
            }
        )
    else:
        st.info("Este proyecto no tiene tareas registradas.")
