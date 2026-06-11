"""
Rex+ Tools — Disponibilidad por Consultor
Selecciona un consultor y muestra sus días ocupados/libres según las tareas
de sus proyectos En curso (filtrados por el campo 'Consultor 1').
"""

import streamlit as st
import requests
import json
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
    hero("📅 Disponibilidad por Consultor", "Selecciona un consultor para ver sus días ocupados y libres según su agenda en Zoho.")
else:
    st.title("📅 Disponibilidad por Consultor")

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

def parse_cf(custom_fields, key):
    for item in custom_fields:
        if key in item:
            v = item[key]
            if isinstance(v, str) and v.startswith("["):
                try:
                    parsed = json.loads(v)
                    return parsed[0] if parsed else ""
                except Exception:
                    return v
            return v
    return ""

@st.cache_data(ttl=900, show_spinner=False)
def get_proyectos(access_token, portal_id):
    """Trae proyectos En curso + BENEFICIO, con su Consultor 1."""
    url = f"https://projectsapi.zoho.com/restapi/portal/{portal_id}/projects/"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    proyectos = []
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
                proyectos.append({
                    "id":        p.get("id"),
                    "key":       p.get("key", ""),
                    "name":      p.get("name", ""),
                    "consultor": parse_cf(p.get("custom_fields", []), "Consultor 1"),
                })
        if len(batch) < 100:
            break
        index += 100
    return proyectos

@st.cache_data(ttl=900, show_spinner=False)
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

# ── Cargar token y proyectos ──────────────────────────────────────────────────

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

with st.spinner("Cargando proyectos..."):
    proyectos = get_proyectos(token, portal_id)

# Lista de consultores disponibles
consultores = sorted({p["consultor"] for p in proyectos if p["consultor"]})

if not consultores:
    st.warning("No se encontraron consultores con proyectos En curso.")
    st.stop()

# ── Selección de consultor ────────────────────────────────────────────────────
st.divider()
c1, c2 = st.columns([2, 1])
with c1:
    consultor_sel = st.selectbox("👤 Selecciona un consultor", ["— Selecciona —"] + consultores)
with c2:
    rango = st.selectbox("📅 Rango", ["Esta semana", "Próxima semana", "Este mes", "Próximos 30 días"])

if consultor_sel == "— Selecciona —":
    st.info("Selecciona un consultor para ver su disponibilidad.")
    st.stop()

# Rango de fechas
hoy = date.today()
if rango == "Esta semana":
    ini = hoy - timedelta(days=hoy.weekday()); fin = ini + timedelta(days=6)
elif rango == "Próxima semana":
    ini = hoy - timedelta(days=hoy.weekday()) + timedelta(days=7); fin = ini + timedelta(days=6)
elif rango == "Este mes":
    ini = date(hoy.year, hoy.month, 1); fin = date(hoy.year, hoy.month, monthrange(hoy.year, hoy.month)[1])
else:
    ini = hoy; fin = hoy + timedelta(days=30)

# Proyectos del consultor
proyectos_consultor = [p for p in proyectos if p["consultor"] == consultor_sel]
st.caption(f"{len(proyectos_consultor)} proyectos · del {ini.strftime('%d/%m/%Y')} al {fin.strftime('%d/%m/%Y')}")

# Ventana ampliada: desde hoy hasta 30 días adelante (para buscar disponibilidad futura)
busqueda_ini = min(ini, hoy)
busqueda_fin = max(fin, hoy + timedelta(days=30))

# Traer tareas de sus proyectos (una sola vez, ventana amplia)
agenda = []       # dentro del rango visible
agenda_full = []  # toda la ventana de búsqueda (para próxima disponibilidad)
barra = st.progress(0.0, text="Cargando agenda...")
for i, p in enumerate(proyectos_consultor):
    tasks = get_tareas_proyecto(token, portal_id, p["id"])
    for t in tasks:
        f_ini = parse_fecha(t.get("start_date_format", "") or t.get("start_date", ""))
        f_fin = parse_fecha(t.get("end_date_format", "") or t.get("end_date", ""))
        if f_ini is None:
            continue
        if f_fin is None or f_fin < f_ini:
            f_fin = f_ini
        d = f_ini
        while d <= f_fin:
            if d.weekday() < 5:
                if ini <= d <= fin:
                    agenda.append({"fecha": d, "tarea": t.get("name", ""), "proyecto": p["name"]})
                if busqueda_ini <= d <= busqueda_fin:
                    agenda_full.append({"fecha": d})
            d += timedelta(days=1)
    barra.progress((i + 1) / max(len(proyectos_consultor), 1), text=f"Cargando agenda... {i+1}/{len(proyectos_consultor)}")
barra.empty()

df_ag = pd.DataFrame(agenda)
dias_ocupados_full = set(pd.DataFrame(agenda_full)["fecha"].unique()) if agenda_full else set()

# Días hábiles del rango
dias_rango = []
d = ini
while d <= fin:
    if d.weekday() < 5:
        dias_rango.append(d)
    d += timedelta(days=1)

dias_ocupados = set(df_ag["fecha"].unique()) if not df_ag.empty else set()
n_ocupados = len([d for d in dias_rango if d in dias_ocupados])
n_libres   = len(dias_rango) - n_ocupados
n_sesiones = len(df_ag.drop_duplicates(subset=["tarea", "proyecto", "fecha"])) if not df_ag.empty else 0

# ── Resumen ───────────────────────────────────────────────────────────────────
st.divider()
k1, k2, k3, k4 = st.columns(4)
k1.metric("Sesiones", n_sesiones)
k2.metric("Días ocupados", n_ocupados)
k3.metric("Días libres", n_libres)
k4.metric("% ocupación", f"{round(n_ocupados/len(dias_rango)*100) if dias_rango else 0}%")

# ── Próxima disponibilidad (mirando 30 días adelante) ─────────────────────────
st.divider()
st.subheader("🔎 Próxima disponibilidad")

DIAS_COMPLETO = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

# Próximo día hábil libre desde mañana
proximo_libre = None
d = hoy + timedelta(days=1)
limite = hoy + timedelta(days=30)
while d <= limite:
    if d.weekday() < 5 and d not in dias_ocupados_full:
        proximo_libre = d
        break
    d += timedelta(days=1)

# Próxima semana (Lun-Vie) completamente despejada
proxima_semana = None
# arrancar el próximo lunes
dlun = hoy + timedelta(days=(7 - hoy.weekday()) % 7 or 7)
while dlun <= limite:
    dias_sem = [dlun + timedelta(days=k) for k in range(5)]
    if all(ds not in dias_ocupados_full for ds in dias_sem):
        proxima_semana = (dias_sem[0], dias_sem[4])
        break
    dlun += timedelta(days=7)

cp1, cp2 = st.columns(2)
with cp1:
    if proximo_libre:
        st.success(f"📆 **Próximo día libre:**\n\n{DIAS_COMPLETO[proximo_libre.weekday()]} {proximo_libre.strftime('%d/%m/%Y')}")
    else:
        st.warning("Sin días libres en los próximos 30 días.")
with cp2:
    if proxima_semana:
        a, b = proxima_semana
        st.success(f"🗓️ **Próxima semana despejada:**\n\n{a.strftime('%d/%m')} al {b.strftime('%d/%m/%Y')}")
    else:
        st.info("Ninguna semana completamente libre en los próximos 30 días.")

# ── Detalle día por día ───────────────────────────────────────────────────────
st.divider()
st.subheader(f"🗓️ Agenda de {consultor_sel}")

DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
for d in dias_rango:
    label = f"{DIAS[d.weekday()]} {d.strftime('%d/%m')}"
    if d in dias_ocupados:
        sesiones = df_ag[df_ag["fecha"] == d].drop_duplicates(subset=["tarea", "proyecto"])
        with st.expander(f"🔴 {label} — {len(sesiones)} sesión(es)", expanded=False):
            for _, s in sesiones.iterrows():
                st.markdown(f"• **{s['tarea']}** · {s['proyecto']}")
    else:
        st.markdown(f"🟢 {label} — *disponible*")

if BRANDING:
    aplicar_footer()
