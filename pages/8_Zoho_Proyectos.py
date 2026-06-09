import streamlit as st
import requests
import pandas as pd
import json
from datetime import datetime

st.set_page_config(page_title="Zoho Proyectos", page_icon="📋", layout="wide")

# ── AUTENTICACIÓN ─────────────────────────────────────────────────────────────

@st.cache_data(ttl=3000, show_spinner=False)
def get_access_token(refresh_token, client_id, client_secret):
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
    url = f"https://projectsapi.zoho.com/restapi/portal/{portal_id}/projects/"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    params = {"status": "active", "range": 100}
    r = requests.get(url, headers=headers, params=params)
    return r.json().get("projects", [])


@st.cache_data(ttl=600, show_spinner=False)
def get_tasks(access_token, portal_id, project_id):
    url = f"https://projectsapi.zoho.com/restapi/portal/{portal_id}/projects/{project_id}/tasks/"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    params = {"range": 100}
    r = requests.get(url, headers=headers, params=params)
    try:
        return r.json().get("tasks", [])
    except Exception:
        return []


# ── HELPERS ───────────────────────────────────────────────────────────────────

def parse_custom_fields(custom_fields):
    """Convierte la lista de dicts custom_fields en un dict plano."""
    result = {}
    if not isinstance(custom_fields, list):
        return result
    for item in custom_fields:
        if isinstance(item, dict):
            for k, v in item.items():
                result[k] = v
    return result


def cf(fields, *keys):
    """Busca un valor en custom_fields probando múltiples keys posibles."""
    for k in keys:
        if k in fields and fields[k] not in (None, "", "false", False):
            val = fields[k]
            # Si es JSON array, parsearlo
            if isinstance(val, str) and val.startswith("["):
                try:
                    parsed = json.loads(val)
                    return ", ".join(parsed) if isinstance(parsed, list) else val
                except Exception:
                    pass
            return val
    return "–"


def status_icon(name):
    name_lower = (name or "").lower()
    if "inicio sin agenda" in name_lower: return "🟠"
    if "en curso"          in name_lower: return "🔵"
    if "ko"                in name_lower: return "🟣"
    if "completado"        in name_lower: return "🟢"
    if "cerrado"           in name_lower: return "⚫"
    return "⚪"


def fmt_date(d):
    if not d or d == "–":
        return "–"
    for fmt in ("%m-%d-%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(d, fmt).strftime("%d/%m/%Y")
        except Exception:
            pass
    return d


def build_df(projects):
    rows = []
    for p in projects:
        cfields = parse_custom_fields(p.get("custom_fields", []))
        status_name = p.get("custom_status_name") or p.get("status", "–")
        task_count  = p.get("task_count", {}) or {}
        rows.append({
            "Clave":        p.get("key", ""),
            "Proyecto":     p.get("name", ""),
            "Estado":       status_icon(status_name) + " " + status_name,
            "Jefe de Proyecto": p.get("owner_name", "–"),
            "Plan":         cf(cfields, "Plan Contratado", "plan_contratado"),
            "Consultor":        cf(cfields, "Consultor 1"),
            "Fecha Facturación": fmt_date(cf(cfields, "Fecha Facturación", "Fecha Facturacion")),
            "Empleados":    cf(cfields, "Cantidad de empleados", "cantidad_de_empleados"),
            "Vendedor":     cf(cfields, "Vendedor", "vendedor"),
            "Razón Social": cf(cfields, "Razón social", "razon_social"),
            "Grupo":        p.get("group_name", "–"),
            "Tareas ✅":    task_count.get("closed", 0),
            "Tareas 🔓":    task_count.get("open", 0),
            "Creado":       fmt_date(p.get("created_date", "")),
            "_id":          str(p.get("id", "")),
            "_nombre_raw":  p.get("name", ""),
            "_status_raw":  status_name,
            "_plan_raw":    cf(cfields, "Plan Contratado", "plan_contratado"),
            "_owner_raw":   p.get("owner_name", ""),
            "_consultor_raw": cf(parse_custom_fields(p.get("custom_fields", [])), "Consultor 1"),
            "_group_raw":   p.get("group_name", ""),
            "_cfields":     cfields,
        })
    return pd.DataFrame(rows)


# ── UI ────────────────────────────────────────────────────────────────────────

st.title("📋 Zoho Proyectos")
st.caption("Proyectos activos del portal Rex+ · datos en tiempo real")

portal_id = st.secrets.get("ZOHO_PORTAL_ID", "757079135")

_, col_btn = st.columns([6, 1])
with col_btn:
    if st.button("🔄 Actualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# Token y proyectos
with st.spinner("Conectando con Zoho Projects..."):
    token = get_access_token(
        st.secrets["ZOHO_REFRESH_TOKEN"],
        st.secrets["ZOHO_CLIENT_ID"],
        st.secrets["ZOHO_CLIENT_SECRET"],
    )

if not token:
    st.error("❌ No se pudo obtener el token. Revisa los secrets en Streamlit.")
    st.stop()

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
k1.metric("Total proyectos",  total)
k2.metric("Sin agendar 🟠",   int(sin_agenda))
k3.metric("En curso 🔵",      int(en_curso))
k4.metric("Otras etapas",     int(otros))
k5.metric("Tareas abiertas",  t_abiertas)

st.divider()

# ── FILTROS ───────────────────────────────────────────────────────────────────
st.subheader("🔍 Filtros")
fc1, fc2, fc3, fc4, fc5 = st.columns(5)

with fc1:
    search = st.text_input("Buscar cliente o proyecto", placeholder="Nombre, razón social...")
with fc2:
    estados_opts = sorted(df["_status_raw"].dropna().unique().tolist())
    filtro_estado = st.multiselect("Estado", estados_opts, placeholder="Todos")
with fc3:
    planes = ["Todos"] + sorted([x for x in df["_plan_raw"].dropna().unique() if x and x != "–"])
    filtro_plan = st.selectbox("Plan", planes)
with fc4:
    consultores = ["Todos"] + sorted([x for x in df["_consultor_raw"].dropna().unique() if x and x != "–"])
    filtro_consultor = st.selectbox("Consultor", consultores)
with fc5:
    grupos_opts = sorted([x for x in df["_group_raw"].dropna().unique() if x])
    filtro_grupo = st.multiselect("Grupo", grupos_opts, placeholder="Todos")

mask = pd.Series([True] * len(df))
if search:
    mask &= (
        df["_nombre_raw"].str.contains(search, case=False, na=False) |
        df["Razón Social"].str.contains(search, case=False, na=False)
    )
if filtro_estado:
    mask &= df["_status_raw"].isin(filtro_estado)
if filtro_plan != "Todos":
    mask &= df["_plan_raw"] == filtro_plan
if filtro_consultor != "Todos":
    mask &= df["_consultor_raw"] == filtro_consultor
if filtro_grupo:
    mask &= df["_group_raw"].isin(filtro_grupo)

df_filtered = df[mask].reset_index(drop=True)
st.caption(f"{len(df_filtered)} proyectos encontrados")

# ── TABLA ─────────────────────────────────────────────────────────────────────
st.divider()
st.subheader("📊 Listado de proyectos")

cols_show = ["Clave", "Proyecto", "Grupo", "Estado", "Consultor", "Jefe de Proyecto", "Plan", "Vendedor", "Empleados", "Fecha Facturación", "Tareas ✅", "Tareas 🔓", "Creado"]
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
    row      = df_filtered[df_filtered["Proyecto"] == selected].iloc[0]
    project_id = row["_id"]
    cfields  = row["_cfields"]

    d1, d2, d3 = st.columns(3)
    with d1:
        st.markdown("**Razón social**")
        st.write(cf(cfields, "Razón social"))
        st.markdown("**RUT**")
        st.write(cf(cfields, "RUT Empresa", "rut_empresa"))
        st.markdown("**Contacto**")
        st.write(cf(cfields, "Jefe de Proyecto Cliente (Contacto)", "nombre_del_contacto"))
        st.markdown("**Correo**")
        st.write(cf(cfields, "Correo del contacto"))
    with d2:
        st.markdown("**Plan contratado**")
        st.write(cf(cfields, "Plan Contratado"))
        st.markdown("**Módulos vendidos**")
        st.write(cf(cfields, "Modulo Vendido"))
        st.markdown("**Vendedor**")
        st.write(cf(cfields, "Vendedor"))
        st.markdown("**Empresa venta**")
        st.write(cf(cfields, "Empresa Venta"))
    with d3:
        st.markdown("**Empleados**")
        st.write(cf(cfields, "Cantidad de empleados"))
        st.markdown("**Empresas**")
        st.write(cf(cfields, "Cantidad de empresas"))
        st.markdown("**Fecha venta**")
        st.write(cf(cfields, "Fecha de Venta"))
        st.markdown("**Teléfono**")
        st.write(cf(cfields, "Telefono de contacto"))

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
            owners = (t.get("details") or {}).get("owners", [])
            responsable = owners[0].get("name", "–") if owners else "–"
            task_rows.append({
                "Tarea":       t.get("name", ""),
                "Estado":      status_t,
                "% avance":    int(t.get("percent_complete", 0) or 0),
                "Responsable": responsable,
                "Vencimiento": fmt_date(t.get("end_date", "")),
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
