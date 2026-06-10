"""
Rex+ Tools — Disponibilidad por Consultor
Muestra días ocupados/libres de cada consultor según sus tareas agendadas en Zoho.
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
def get_all_tasks(access_token, portal_id):
    """Trae todas las tareas del portal vía mytasks?tasktype=all, paginado."""
    url = f"https://projectsapi.zoho.com/restapi/portal/{portal_id}/mytasks/"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    all_tasks = []
    index = 1
    while True:
        r = requests.get(url, headers=headers, params={"tasktype": "all", "range": 200, "index": index})
        if r.status_code != 200:
            break
        batch = r.json().get("tasks", [])
        if not batch:
            break
        all_tasks.extend(batch)
        if len(batch) < 200:
            break
        index += 200
        if index > 5000:  # tope de seguridad
            break
    return all_tasks

def parse_fecha(s):
    """Parsea '06-09-2026 11:00:00 AM' o '06-09-2026' a date."""
    if not s:
        return None
    s = s.split(" ")[0]  # quedarse solo con la fecha
    for fmt in ("%m-%d-%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None

def build_agenda(tasks):
    """Construye DataFrame: consultor, fecha, tarea, proyecto."""
    rows = []
    for t in tasks:
        owners = (t.get("details") or {}).get("owners", [])
        fecha  = parse_fecha(t.get("start_date_format", "") or t.get("start_date", ""))
        if fecha is None:
            continue
        proyecto = (t.get("project") or {}).get("name", "")
        for o in owners:
            nombre = o.get("full_name") or o.get("name", "")
            if not nombre or nombre.lower() in ("sin asignar", "unassigned"):
                continue
            rows.append({
                "consultor": nombre,
                "fecha":     fecha,
                "tarea":     t.get("name", ""),
                "proyecto":  proyecto,
                "completada": t.get("completed", False),
            })
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

with st.spinner("Cargando tareas agendadas..."):
    tasks = get_all_tasks(token, portal_id)

df = build_agenda(tasks)

if df.empty:
    st.warning("No se encontraron tareas con fecha asignadas a consultores.")
    st.stop()

# ── Filtros de rango ──────────────────────────────────────────────────────────
st.divider()
hoy = date.today()
fc1, fc2, fc3 = st.columns(3)
with fc1:
    rango = st.selectbox("Rango", ["Esta semana", "Próxima semana", "Este mes", "Personalizado"])
with fc2:
    if rango == "Personalizado":
        f_desde = st.date_input("Desde", value=hoy)
    else:
        f_desde = None
with fc3:
    if rango == "Personalizado":
        f_hasta = st.date_input("Hasta", value=hoy + timedelta(days=7))
    else:
        f_hasta = None

# Calcular rango de fechas
if rango == "Esta semana":
    ini = hoy - timedelta(days=hoy.weekday())
    fin = ini + timedelta(days=6)
elif rango == "Próxima semana":
    ini = hoy - timedelta(days=hoy.weekday()) + timedelta(days=7)
    fin = ini + timedelta(days=6)
elif rango == "Este mes":
    ini = date(hoy.year, hoy.month, 1)
    fin = date(hoy.year, hoy.month, monthrange(hoy.year, hoy.month)[1])
else:
    ini, fin = f_desde, f_hasta

st.caption(f"Mostrando del {ini.strftime('%d/%m/%Y')} al {fin.strftime('%d/%m/%Y')}")

# Días hábiles del rango (Lun-Vie, ya que weekends son Vie/Sab según portal — usamos Lun-Jue+Vie)
dias_rango = []
d = ini
while d <= fin:
    if d.weekday() < 5:  # 0=Lun ... 4=Vie
        dias_rango.append(d)
    d += timedelta(days=1)

# Filtrar agenda al rango
df_rango = df[(df["fecha"] >= ini) & (df["fecha"] <= fin)]

# ── TABLA RESUMEN: ocupado/libre por consultor ────────────────────────────────
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
        "Consultor":      c,
        "Sesiones":       len(tareas_c),
        "Días ocupados":  n_ocupados,
        "Días libres":    n_libres,
        "% ocupación":    round(n_ocupados / len(dias_rango) * 100) if dias_rango else 0,
    })

df_resumen = pd.DataFrame(resumen).sort_values("Días libres", ascending=False)
st.dataframe(
    df_resumen,
    use_container_width=True,
    hide_index=True,
    column_config={
        "% ocupación": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%d%%"),
        "Días libres": st.column_config.NumberColumn(width="small"),
        "Días ocupados": st.column_config.NumberColumn(width="small"),
    }
)

# ── DETALLE POR CONSULTOR ─────────────────────────────────────────────────────
st.divider()
st.subheader("🗓️ Detalle de agenda")

consultor_sel = st.selectbox("Selecciona un consultor", consultores)

tareas_sel = df_rango[df_rango["consultor"] == consultor_sel].sort_values("fecha")
dias_ocupados_sel = set(tareas_sel["fecha"].unique())

# Grilla de días del rango
st.markdown(f"**Agenda de {consultor_sel}**")

for d in dias_rango:
    dia_nombre = ["Lunes","Martes","Miércoles","Jueves","Viernes"][d.weekday()]
    label = f"{dia_nombre} {d.strftime('%d/%m')}"
    if d in dias_ocupados_sel:
        sesiones = tareas_sel[tareas_sel["fecha"] == d]
        with st.expander(f"🔴 {label} — {len(sesiones)} sesión(es)", expanded=False):
            for _, s in sesiones.iterrows():
                st.markdown(f"• **{s['tarea']}** · {s['proyecto']}")
    else:
        st.markdown(f"🟢 {label} — *disponible*")

if BRANDING:
    aplicar_footer()
