"""
Rex+ Tools — Disponibilidad por Consultor
Días ocupados/libres de cada consultor según tareas agendadas en proyectos
En curso + BENEFICIO EMPRESA REXMAS.
"""

import streamlit as st
import requests
import pandas as pd
from datetime import datetime, date, timedelta
from calendar import monthrange

try:
    from lib.branding import aplicar_branding, aplicar_footer, hero
    BRANDING = True
except ImportError:
    BRANDING = False

st.set_page_config(page_title="Disponibilidad | Rex+ Tools", page_icon="📅", layout="wide")

if BRANDING:
    aplicar_branding(titulo_pagina="Disponibilidad", badge="PRODUCCIÓN")
    hero("📅 Disponibilidad por Consultor", "Días ocupados y libres según las sesiones agendadas en Zoho Projects.")
else:
    st.title("📅 Disponibilidad por Consultor")
    st.caption("Días ocupados y libres según las sesiones agendadas en Zoho Projects.")

# ── ZOHO ──────────────────────────────────────────────────────────────────────

@st.cache_data(ttl=3000, show_spinner=False)
def get_access_token(refresh_token, client_id, client_secret):
    r = requests.post("https://accounts.zoho.com/oauth/v2/token", params={
        "refresh_token": refresh_token,
        "client_id":     client_id,
        "client_secret": client_secret,
        "grant_type":    "refresh_token",
    })
    return r.json().get("access_token")

@st.cache_data(ttl=600, show_spinner=False)
def get_proyectos_relevantes(access_token, portal_id):
    """Proyectos En curso + BENEFICIO EMPRESA REXMAS."""
    url = f"https://projectsapi.zoho.com/restapi/portal/{portal_id}/projects/"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    relevantes = []
    index = 1
    while True:
        r = requests.get(url, headers=headers, params={"range": 100, "index": index})
        batch = r.json().get("projects", [])
        if not batch:
            break
        for p in batch:
            status = p.get("custom_status_name", "").lower()
            nombre = p.get("name", "").upper()
            if status == "en curso" or "BENEFICIO EMPRESA REXMAS" in nombre:
                relevantes.append({"id": p.get("id"), "key": p.get("key", ""), "name": p.get("name", "")})
        if len(batch) < 100:
            break
        index += 100
    return relevantes

@st.cache_data(ttl=600, show_spinner=False)
def get_tareas_proyecto(access_token, portal_id, project_id):
    url = f"https://projectsapi.zoho.com/restapi/portal/{portal_id}/projects/{project_id}/tasks/"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    r = requests.get(url, headers=headers, params={"range": 200})
    if r.status_code != 200:
        return []
    return r.json().get("tasks", [])

def parse_fecha(s):
    if not s:
        return None
    s = str(s).split(" ")[0]
    for fmt in ("%m-%d-%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None

@st.cache_data(ttl=600, show_spinner=True)
def construir_agenda(_token, portal_id):
    """Recorre proyectos relevantes y arma DataFrame de agenda."""
    proyectos = get_proyectos_relevantes(_token, portal_id)
    rows = []
    barra = st.progress(0.0, text="Cargando agendas...")
    total = len(proyectos)
    for i, p in enumerate(proyectos):
        tasks = get_tareas_proyecto(_token, portal_id, p["id"])
        for t in tasks:
            fecha = parse_fecha(t.get("start_date_format", "") or t.get("start_date", ""))
            if fecha is None:
                continue
            owners = (t.get("details") or {}).get("owners", [])
            for o in owners:
                nombre = o.get("full_name") or o.get("name", "")
                if not nombre or nombre.lower() in ("sin asignar", "unassigned"):
                    continue
                rows.append({
                    "consultor": nombre,
                    "fecha":     fecha,
                    "tarea":     t.get("name", ""),
                    "proyecto":  p["name"],
                    "completada": t.get("completed", False),
                })
        barra.progress((i + 1) / total, text=f"Cargando agendas... {i+1}/{total}")
    barra.empty()
    return pd.DataFrame(rows)

# ── Cargar datos ──────────────────────────────────────────────────────────────

portal_id = st.secrets.get("ZOHO_PORTAL_ID", "757079135")

_, col_btn = st.columns([6, 1])
with col_btn:
    if st.button("🔄 Actualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

try:
    token = get_access_token(
        st.secrets["ZOHO_REFRESH_TOKEN"],
        st.secrets["ZOHO_CLIENT_ID"],
        st.secrets["ZOHO_CLIENT_SECRET"],
    )
    ZOHO_OK = bool(token)
except Exception:
    token = None
    ZOHO_OK = False

if not ZOHO_OK:
    st.error("❌ No se pudo conectar con Zoho.")
    st.stop()

df = construir_agenda(token, portal_id)

if df.empty:
    st.warning("No se encontraron tareas con fecha asignadas a consultores.")
    st.stop()

# ── Filtros de rango ──────────────────────────────────────────────────────────
st.divider()
hoy = date.today()
fc1, fc2, fc3 = st.columns(3)
with fc1:
    rango = st.selectbox("Rango", ["Esta semana", "Próxima semana", "Este mes", "Próximos 30 días", "Personalizado"])
with fc2:
    f_desde = st.date_input("Desde", value=hoy, key="f_desde") if rango == "Personalizado" else None
with fc3:
    f_hasta = st.date_input("Hasta", value=hoy + timedelta(days=7), key="f_hasta") if rango == "Personalizado" else None

if rango == "Esta semana":
    ini = hoy - timedelta(days=hoy.weekday()); fin = ini + timedelta(days=6)
elif rango == "Próxima semana":
    ini = hoy - timedelta(days=hoy.weekday()) + timedelta(days=7); fin = ini + timedelta(days=6)
elif rango == "Este mes":
    ini = date(hoy.year, hoy.month, 1); fin = date(hoy.year, hoy.month, monthrange(hoy.year, hoy.month)[1])
elif rango == "Próximos 30 días":
    ini = hoy; fin = hoy + timedelta(days=30)
else:
    ini, fin = f_desde, f_hasta

st.caption(f"Mostrando del {ini.strftime('%d/%m/%Y')} al {fin.strftime('%d/%m/%Y')}")

# Días hábiles (Lun-Vie)
dias_rango = []
d = ini
while d <= fin:
    if d.weekday() < 5:
        dias_rango.append(d)
    d += timedelta(days=1)

df_rango = df[(df["fecha"] >= ini) & (df["fecha"] <= fin)]

# ── TABLA RESUMEN ─────────────────────────────────────────────────────────────
st.divider()
st.subheader("📊 Resumen de disponibilidad")

consultores = sorted(df["consultor"].unique())
resumen = []
for c in consultores:
    tareas_c = df_rango[df_rango["consultor"] == c]
    dias_ocupados = set(tareas_c["fecha"].unique())
    n_ocupados = len([d for d in dias_rango if d in dias_ocupados])
    n_libres   = len(dias_rango) - n_ocupados
    resumen.append({
        "Consultor":     c,
        "Sesiones":      len(tareas_c),
        "Días ocupados": n_ocupados,
        "Días libres":   n_libres,
        "% ocupación":   round(n_ocupados / len(dias_rango) * 100) if dias_rango else 0,
    })

df_resumen = pd.DataFrame(resumen).sort_values("Días libres", ascending=False)
st.dataframe(
    df_resumen,
    use_container_width=True,
    hide_index=True,
    column_config={
        "% ocupación":   st.column_config.ProgressColumn(min_value=0, max_value=100, format="%d%%"),
        "Días libres":   st.column_config.NumberColumn(width="small"),
        "Días ocupados": st.column_config.NumberColumn(width="small"),
    }
)

# ── DETALLE POR CONSULTOR ─────────────────────────────────────────────────────
st.divider()
st.subheader("🗓️ Detalle de agenda")

consultor_sel = st.selectbox("Selecciona un consultor", consultores)
tareas_sel = df_rango[df_rango["consultor"] == consultor_sel].sort_values("fecha")
dias_ocupados_sel = set(tareas_sel["fecha"].unique())

st.markdown(f"**Agenda de {consultor_sel}**")
DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
for d in dias_rango:
    label = f"{DIAS[d.weekday()]} {d.strftime('%d/%m')}"
    if d in dias_ocupados_sel:
        sesiones = tareas_sel[tareas_sel["fecha"] == d]
        with st.expander(f"🔴 {label} — {len(sesiones)} sesión(es)", expanded=False):
            for _, s in sesiones.iterrows():
                st.markdown(f"• **{s['tarea']}** · {s['proyecto']}")
    else:
        st.markdown(f"🟢 {label} — *disponible*")

if BRANDING:
    aplicar_footer()
