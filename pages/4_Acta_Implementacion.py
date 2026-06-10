"""
Acta de Implementación REX+
P�gina integrada al hub rex-tools usando el branding compartido.
"""

import sys
import json
import datetime
import requests
from pathlib import Path

import streamlit as st
import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
_ROOT     = Path(__file__).parent.parent
_ACTA_DIR = _ROOT / "acta_app"
_LIB_DIR  = _ROOT / "lib"

sys.path.insert(0, str(_ACTA_DIR))
sys.path.insert(0, str(_LIB_DIR))

from branding import aplicar_branding, aplicar_footer, hero
from extractor import extract_all
from generator import generate_acta

st.set_page_config(
    page_title="Acta de Implementación · REX+",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

aplicar_branding(titulo_pagina="Acta de Implementación", badge="IMPLEMENTACIÓN")
hero(
    titulo="Acta de Implementación",
    descripcion="Genera actas automáticamente a partir de los datos del cliente y las notas de sesión de Gemini.",
    icono="📄",
)

CLIENTES_PATH = _ACTA_DIR / "clientes.json"
EQUIPO_PATH   = _ACTA_DIR / "equipo.json"

# ── Helpers ───────────────────────────────────────────────────────────────────

def load_clients():
    if CLIENTES_PATH.exists():
        with open(CLIENTES_PATH, encoding='utf-8') as f:
            return json.load(f)
    return {"clientes": []}

def save_client(client: dict):
    data = load_clients()
    idx = next((i for i, c in enumerate(data["clientes"])
                if c["empresa"].lower() == client["empresa"].lower()), None)
    if idx is not None:
        data["clientes"][idx] = client
    else:
        data["clientes"].append(client)
    with open(CLIENTES_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def delete_client(empresa: str):
    data = load_clients()
    data["clientes"] = [c for c in data["clientes"] if c["empresa"] != empresa]
    with open(CLIENTES_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def load_equipo() -> dict:
    if EQUIPO_PATH.exists():
        with open(EQUIPO_PATH, encoding='utf-8') as f:
            return json.load(f)
    return {"consultores": [], "jefes": []}

def save_equipo(data: dict):
    with open(EQUIPO_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def add_equipo_member(rol: str, nombre: str, email: str, telefono: str = ""):
    data = load_equipo()
    if not any(m["nombre"].lower() == nombre.lower() for m in data[rol]):
        data[rol].append({"nombre": nombre, "email": email, "telefono": telefono})
        save_equipo(data)

def update_equipo_telefono(rol: str, nombre: str, telefono: str):
    data = load_equipo()
    for m in data[rol]:
        if m["nombre"] == nombre:
            m["telefono"] = telefono
            break
    save_equipo(data)

def delete_equipo_member(rol: str, nombre: str):
    data = load_equipo()
    data[rol] = [m for m in data[rol] if m["nombre"] != nombre]
    save_equipo(data)

def reset_notes():
    for k in ["dev_points", "activities", "notes_extracted", "extracted_fecha"]:
        v = st.session_state[k]
        st.session_state[k] = ([] if isinstance(v, list)
                               else False if isinstance(v, bool)
                               else "")

def step_pill(num: int, label: str, done: bool = False):
    icon = "✓" if done else str(num)
    bg   = "#1EBBEF" if done else "#1A3A5F"
    st.markdown(f"""
    <div style="display:inline-flex;align-items:center;gap:8px;
                background:white;border:1px solid #dde3f0;
                border-radius:99px;padding:5px 14px 5px 6px;
                margin-bottom:10px;box-shadow:0 1px 3px rgba(0,0,0,0.06);">
        <span style="background:{bg};color:white;border-radius:50%;
                     width:22px;height:22px;display:inline-flex;
                     align-items:center;justify-content:center;
                     font-size:11px;font-weight:700;">{icon}</span>
        <span style="font-size:12px;font-weight:600;color:#1A3A5F;">{label}</span>
    </div>""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
for k, v in {
    "dev_points": [], "activities": [],
    "notes_extracted": False, "extracted_fecha": "",
    "docx_bytes": None, "docx_filename": "",
    "edit_mode": False, "last_selected": "",
    "drive_files": None,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── ZOHO HELPERS ─────────────────────────────────────────────────────────────

@st.cache_data(ttl=3000, show_spinner=False)
def _get_token(refresh_token, client_id, client_secret):
    r = requests.post("https://accounts.zoho.com/oauth/v2/token", params={
        "refresh_token": refresh_token,
        "client_id":     client_id,
        "client_secret": client_secret,
        "grant_type":    "refresh_token",
    })
    return r.json().get("access_token")

@st.cache_data(ttl=600, show_spinner=False)
def _listar_ots(access_token, portal_id):
    ESTADOS = ["inicio sin agenda", "reunion ko", "reunión ko", "agenda por confirmar"]
    url = f"https://projectsapi.zoho.com/restapi/portal/{portal_id}/projects/"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    rows = []
    index = 1
    while True:
        r = requests.get(url, headers=headers, params={"range": 100, "index": index})
        batch = r.json().get("projects", [])
        if not batch:
            break
        for p in batch:
            status = p.get("custom_status_name", "")
            if status.lower() in ESTADOS:
                cfields = _parse_cf(p.get("custom_fields", []))
                rows.append({
                    "OT":        p.get("key", ""),
                    "Proyecto":  p.get("name", ""),
                    "Consultor": _cf(cfields, "Consultor 1"),
                    "Estado":    status,
                })
        if len(batch) < 100:
            break
        index += 100
    return rows

@st.cache_data(ttl=600, show_spinner=False)
def _buscar_ot(access_token, portal_id, ot):
    url = f"https://projectsapi.zoho.com/restapi/portal/{portal_id}/projects/"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    index = 1
    while True:
        r = requests.get(url, headers=headers, params={"range": 100, "index": index})
        batch = r.json().get("projects", [])
        if not batch:
            break
        for p in batch:
            ot_up = ot.strip().upper()
            if ot_up == p.get("key", "").upper() or ot_up in p.get("name", "").upper():
                return p
        if len(batch) < 100:
            break
        index += 100
    return None

def _parse_cf(custom_fields):
    result = {}
    if isinstance(custom_fields, list):
        for item in custom_fields:
            if isinstance(item, dict):
                for k, v in item.items():
                    result[k] = v
    return result

def _cf(fields, *keys):
    for k in keys:
        if k in fields and fields[k] not in (None, "", "false", False):
            return str(fields[k])
    return ""

def _extraer_zoho(proyecto):
    if not proyecto:
        return {}
    cfields = _parse_cf(proyecto.get("custom_fields", []))
    return {
        "empresa":    _cf(cfields, "Razón social"),
        "jefe":       _cf(cfields, "Jefe de Proyecto Cliente (Contacto)"),
        "correo":     _cf(cfields, "Correo del contacto"),
        "telefono":   _cf(cfields, "Telefono de contacto"),
        "nombre_proyecto": proyecto.get("name", ""),
    }

try:
    _token = _get_token(
        st.secrets["ZOHO_REFRESH_TOKEN"],
        st.secrets["ZOHO_CLIENT_ID"],
        st.secrets["ZOHO_CLIENT_SECRET"],
    )
    _PORTAL_ID = st.secrets.get("ZOHO_PORTAL_ID", "757079135")
    ZOHO_OK = bool(_token)
except Exception:
    _token = None
    _PORTAL_ID = "757079135"
    ZOHO_OK = False

# Session state para OT
for k, v in {"zoho_acta": {}, "last_ot_acta": "", "zoho_acta_msg": ()}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── Layout ────────────────────────────────────────────────────────────────────
col_form, col_right = st.columns([3, 2], gap="large")

# ═══════════════════════════════════════════════════════════════════════════════
# COLUMNA IZQUIERDA
# ═══════════════════════════════════════════════════════════════════════════════
with col_form:

    # ── BÚSQUEDA POR OT ───────────────────────────────────────────────────────
    st.markdown("### 🔍 Buscar proyecto por OT")

    if ZOHO_OK:
        with st.expander("📋 Ver OTs activas (Inicio sin agenda / Reunión KO / Agenda por confirmar)", expanded=False):
            with st.spinner("Cargando OTs..."):
                ots = _listar_ots(_token, _PORTAL_ID)
            if ots:
                busq = st.text_input("🔍 Filtrar", placeholder="Buscar por OT, nombre o consultor...", key="acta_busq_ot")
                df_ots = pd.DataFrame(ots)
                if busq:
                    mask = df_ots.apply(lambda row: row.astype(str).str.contains(busq, case=False).any(), axis=1)
                    df_ots = df_ots[mask].reset_index(drop=True)
                st.dataframe(df_ots, use_container_width=True, hide_index=True,
                             column_config={"OT": st.column_config.TextColumn(width="small"),
                                            "Proyecto": st.column_config.TextColumn(width="large")},
                             key="acta_tabla_ots")
                st.caption(f"{len(df_ots)} proyectos · copia la OT y pégala en el campo de abajo")
            else:
                st.info("No hay proyectos en estos estados.")

    ot_input = st.text_input("OT (Orden de Trabajo)", placeholder="Ej: RE-2910 o 2910", key="acta_ot")

    if ot_input and ot_input != st.session_state.last_ot_acta and ZOHO_OK:
        with st.spinner(f"Buscando OT {ot_input} en Zoho..."):
            proyecto = _buscar_ot(_token, _PORTAL_ID, ot_input)
        if proyecto:
            datos = _extraer_zoho(proyecto)
            st.session_state.zoho_acta      = datos
            st.session_state.last_ot_acta   = ot_input
            st.session_state.zoho_acta_msg  = ("ok", f"✅ Proyecto encontrado: **{datos['nombre_proyecto']}**")
        else:
            st.session_state.zoho_acta      = {}
            st.session_state.last_ot_acta   = ot_input
            st.session_state.zoho_acta_msg  = ("warn", f"⚠️ No se encontró proyecto con OT **{ot_input}**.")
        st.rerun()

    if st.session_state.zoho_acta_msg:
        tipo, msg = st.session_state.zoho_acta_msg
        if tipo == "ok":
            st.success(msg)
        else:
            st.warning(msg)

    st.divider()

    # ── PASO 1: Cliente ───────────────────────────────────────────────────────
    step_pill(1, "Selecciona o ingresa el cliente",
              done=st.session_state.edit_mode is False and st.session_state.last_selected != "")

    data         = load_clients()
    client_list  = data["clientes"]
    client_names = [c["nombre_display"] for c in client_list]
    options      = ["— Nuevo cliente —"] + client_names

    col_sel, col_edit, col_del = st.columns([4, 1, 1])
    with col_sel:
        selected = st.selectbox("Cliente guardado", options, index=0, label_visibility="collapsed")
    if selected != st.session_state.last_selected:
        st.session_state.edit_mode    = False
        st.session_state.last_selected = selected
    with col_edit:
        if selected != "— Nuevo cliente —":
            label = "🔓 Editar" if not st.session_state.edit_mode else "🔒 Bloquear"
            if st.button(label, use_container_width=True):
                st.session_state.edit_mode = not st.session_state.edit_mode
                st.rerun()
    with col_del:
        if selected != "— Nuevo cliente —":
            if st.button("🗑", use_container_width=True, help="Eliminar cliente"):
                delete_client(next(c["empresa"] for c in client_list if c["nombre_display"] == selected))
                st.session_state.edit_mode = False
                st.rerun()

    client   = next((c for c in client_list if c["nombre_display"] == selected), None)
    is_new   = selected == "— Nuevo cliente —"
    editable = is_new or st.session_state.edit_mode

    if client and not editable:
        st.info("🔒 Datos en modo lectura — haz clic en **🔓 Editar** para modificar")

    st.divider()

    # ── PASO 2: Datos del cliente ─────────────────────────────────────────────
    step_pill(2, "Datos del cliente")
    st.markdown("### Datos del cliente")

    # Datos desde Zoho para nuevo cliente
    _z = st.session_state.zoho_acta

    empresa      = st.text_input("Empresa *",
                                 value=client["empresa"] if client else _z.get("empresa", ""),
                                 placeholder="RE-XXXX - Nombre empresa (ALIAS)",
                                 disabled=not editable)
    jefe_cliente = st.text_input("Jefe de proyecto cliente *",
                                 value=client["jefe_cliente"] if client else _z.get("jefe", ""),
                                 placeholder="Nombre completo en mayúsculas",
                                 disabled=not editable)
    col1, col2, col3 = st.columns(3)
    with col1:
        email_jefe_cliente = st.text_input("Email jefe cliente",
                                           value=client.get("email_jefe_cliente", "") if client else _z.get("correo", ""),
                                           placeholder="nombre@empresa.cl",
                                           disabled=not editable)
    with col2:
        tel_jefe_cliente = st.text_input("Teléfono jefe cliente",
                                         value=client.get("tel_jefe_cliente", "") if client else _z.get("telefono", ""),
                                         placeholder="+56 9 XXXX XXXX",
                                         disabled=not editable)
    with col3:
        pass

    st.markdown("**Usuario implementador**")
    col1, col2, col3 = st.columns(3)
    with col1:
        usuario_impl = st.text_input("Nombre *",
                                     value=client["usuario_impl"] if client else "",
                                     disabled=not editable)
    with col2:
        email_impl = st.text_input("Email",
                                   value=client["email_impl"] if client else "",
                                   placeholder="nombre@empresa.cl",
                                   disabled=not editable)
    with col3:
        tel_usuario_impl = st.text_input("Teléfono",
                                         value=client.get("tel_usuario_impl", "") if client else "",
                                         placeholder="+56 9 XXXX XXXX",
                                         disabled=not editable)

    st.divider()

    # ── PASO 3: Equipo REX+ ───────────────────────────────────────────────────
    step_pill(3, "Equipo REX+")
    st.markdown("### Equipo REX+")

    equipo              = load_equipo()
    jefes_list          = equipo.get("jefes", [])
    consultores_list    = equipo.get("consultores", [])
    jefes_nombres       = [m["nombre"] for m in jefes_list]
    consultores_nombres = [m["nombre"] for m in consultores_list]

    col1, col2 = st.columns(2)
    with col1:
        jefe_opts    = ["— Nuevo —"] + jefes_nombres
        jefe_default = client["jefe_rex"] if client else (jefes_nombres[0] if jefes_nombres else "— Nuevo —")
        jefe_idx     = jefe_opts.index(jefe_default) if jefe_default in jefe_opts else 0
        jefe_sel     = st.selectbox("Jefe de proyecto REX+", jefe_opts, index=jefe_idx)

        if jefe_sel == "— Nuevo —":
            jefe_rex       = st.text_input("Nombre jefe *", placeholder="NOMBRE APELLIDO")
            email_jefe_rex = st.text_input("Email jefe", placeholder="nombre@visma.com")
            tel_jefe_rex   = st.text_input("Teléfono jefe", placeholder="+56 9 XXXX XXXX", key="tel_jefe_new")
        else:
            jefe_data      = next(m for m in jefes_list if m["nombre"] == jefe_sel)
            jefe_rex       = jefe_sel
            email_jefe_rex = st.text_input("Email jefe", value=jefe_data["email"], key="email_jefe_exist")
            tel_jefe_rex   = st.text_input("Teléfono jefe",
                                           value=jefe_data.get("telefono", ""),
                                           placeholder="+56 9 XXXX XXXX",
                                           key="tel_jefe_exist")
            # Guardar teléfono si se ingresó
            tel_guardado_j = jefe_data.get("telefono", "")
            if tel_jefe_rex and tel_jefe_rex != tel_guardado_j:
                if st.button("💾 Guardar teléfono jefe", key="save_tel_jefe", use_container_width=True):
                    update_equipo_telefono("jefes", jefe_sel, tel_jefe_rex)
                    st.success("✓ Teléfono guardado.")
                    st.rerun()

    with col2:
        cons_opts    = ["— Nuevo —"] + consultores_nombres
        cons_default = client["consultor"] if client else (consultores_nombres[0] if consultores_nombres else "— Nuevo —")
        cons_idx     = cons_opts.index(cons_default) if cons_default in cons_opts else 0
        cons_sel     = st.selectbox("Consultor REX", cons_opts, index=cons_idx)

        if cons_sel == "— Nuevo —":
            consultor       = st.text_input("Nombre consultor *", placeholder="NOMBRE APELLIDO")
            email_consultor = st.text_input("Email consultor", placeholder="nombre@visma.com")
            tel_consultor   = st.text_input("Teléfono consultor", placeholder="+56 9 XXXX XXXX", key="tel_cons_new")
        else:
            cons_data       = next(m for m in consultores_list if m["nombre"] == cons_sel)
            consultor       = cons_sel
            email_consultor = st.text_input("Email consultor", value=cons_data["email"], key="email_cons_exist")
            tel_consultor   = st.text_input("Teléfono consultor",
                                            value=cons_data.get("telefono", ""),
                                            placeholder="+56 9 XXXX XXXX",
                                            key="tel_cons_exist")
            # Guardar teléfono si se ingresó
            tel_guardado_c = cons_data.get("telefono", "")
            if tel_consultor and tel_consultor != tel_guardado_c:
                if st.button("💾 Guardar teléfono consultor", key="save_tel_cons", use_container_width=True):
                    update_equipo_telefono("consultores", cons_sel, tel_consultor)
                    st.success("✓ Teléfono guardado.")
                    st.rerun()

        horas = st.number_input("Horas de sesión", min_value=1, max_value=8,
                                value=int(client["horas"]) if client else 4)

    col_gj, col_gc = st.columns(2)
    with col_gj:
        if jefe_sel == "— Nuevo —" and jefe_rex:
            if st.button("💾 Guardar jefe", use_container_width=True):
                add_equipo_member("jefes", jefe_rex, email_jefe_rex, tel_jefe_rex)
                st.success(f"✓ Jefe '{jefe_rex}' guardado.")
                st.rerun()
    with col_gc:
        if cons_sel == "— Nuevo —" and consultor:
            if st.button("💾 Guardar consultor", use_container_width=True):
                add_equipo_member("consultores", consultor, email_consultor, tel_consultor)
                st.success(f"✓ Consultor '{consultor}' guardado.")
                st.rerun()

    with st.expander("🗑 Gestionar equipo guardado"):
        col_ej, col_ec = st.columns(2)
        with col_ej:
            st.markdown("**Jefes**")
            for m in jefes_list:
                c1, c2 = st.columns([3, 1])
                c1.write(m["nombre"])
                if c2.button("🗑", key=f"del_j_{m['nombre']}"):
                    delete_equipo_member("jefes", m["nombre"]); st.rerun()
        with col_ec:
            st.markdown("**Consultores**")
            for m in consultores_list:
                c1, c2 = st.columns([3, 1])
                c1.write(m["nombre"])
                if c2.button("🗑", key=f"del_c_{m['nombre']}"):
                    delete_equipo_member("consultores", m["nombre"]); st.rerun()

    st.divider()

    # ── PASO 4: Sesión ────────────────────────────────────────────────────────
    step_pill(4, "Datos de la sesión")
    st.markdown("### Datos de la sesión")

    plan_opts = ["Full", "Estándar", "Express", "Base", "Casino con instalación",
                 "Casino sin instalación", "Asistencia solo Marcaje app"]
    plan_idx  = plan_opts.index(client["plan"]) if client and client.get("plan") in plan_opts else 0
    col1, col2, col3 = st.columns(3)
    with col1:
        plan = st.selectbox("Plan", plan_opts, index=plan_idx)
    with col2:
        acta_num = st.text_input("N° de acta",
                                 value=st.session_state.extracted_fecha,
                                 placeholder="Ej: 27-04-2026")
    with col3:
        fecha = st.date_input("Fecha de sesión", value=datetime.date.today())

    st.divider()

    # ── PASO 5: Asistentes ────────────────────────────────────────────────────
    step_pill(5, "Asistentes a la sesión")
    st.markdown("### Asistentes a la sesión")

    # Construir lista base desde el cliente guardado, o desde los campos actuales
    if client and client.get("asistentes"):
        default_asistentes = client["asistentes"]
    else:
        default_asistentes = [
            {"nombre": client["usuario_impl"] if client else usuario_impl, "cargo": "", "gerencia": ""},
            {"nombre": client["jefe_cliente"] if client else jefe_cliente,  "cargo": "", "gerencia": ""},
            {"nombre": jefe_rex,   "cargo": "", "gerencia": ""},
            {"nombre": consultor,  "cargo": "", "gerencia": ""},
        ]

    edited = st.data_editor(
        pd.DataFrame(default_asistentes),
        column_config={
            "nombre":   st.column_config.TextColumn("Nombre",   width="medium"),
            "cargo":    st.column_config.TextColumn("Cargo",    width="medium"),
            "gerencia": st.column_config.TextColumn("Gerencia", width="medium"),
        },
        num_rows="dynamic", use_container_width=True, hide_index=True,
    )

    st.divider()

    # ── Guardar cliente ───────────────────────────────────────────────────────
    def client_payload(nombre_display="", keyword_drive=""):
        asistentes_actuales = [
            {"nombre": str(row["nombre"]), "cargo": str(row.get("cargo", "")),
             "gerencia": str(row.get("gerencia", ""))}
            for _, row in edited.iterrows()
            if str(row["nombre"]).strip()
        ]
        return {
            "nombre_display":     nombre_display or empresa[:20],
            "empresa":            empresa,
            "plan":               plan,
            "jefe_cliente":       jefe_cliente,
            "email_jefe_cliente": email_jefe_cliente,
            "tel_jefe_cliente":   tel_jefe_cliente,
            "usuario_impl":       usuario_impl,
            "email_impl":         email_impl,
            "tel_usuario_impl":   tel_usuario_impl,
            "jefe_rex":           jefe_rex,
            "email_jefe_rex":     email_jefe_rex,
            "consultor":          consultor,
            "horas":              int(horas),
            "keyword_drive":      keyword_drive,
            "asistentes":         asistentes_actuales,
        }

    if is_new:
        with st.expander("💾 Guardar nuevo cliente"):
            nombre_display = st.text_input("Nombre corto del cliente",
                                           value="", placeholder="Ej: CYGNUS")
            if st.button("Guardar cliente", use_container_width=True, type="primary"):
                if not empresa:
                    st.error("Ingresa el nombre de la empresa.")
                else:
                    save_client(client_payload(
                        nombre_display,
                        nombre_display.lower() if nombre_display else "",
                    ))
                    st.success(f"✓ Cliente '{nombre_display}' guardado.")
                    st.session_state.edit_mode = False
                    reset_notes()  # ← CAMBIO 1
                    st.rerun()
    elif st.session_state.edit_mode:
        if st.button("💾 Guardar cambios", use_container_width=True, type="primary"):
            if not empresa:
                st.error("Ingresa el nombre de la empresa.")
            else:
                save_client(client_payload(
                    client["nombre_display"],
                    client.get("keyword_drive", ""),
                ))
                st.success(f"✓ Cambios guardados para **{client['nombre_display']}**.")
                st.session_state.edit_mode = False
                reset_notes()  # ← CAMBIO 1
                st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# COLUMNA DERECHA
# ═══════════════════════════════════════════════════════════════════════════════
with col_right:

    step_pill(6, "Sube las notas de Gemini", done=st.session_state.notes_extracted)
    st.markdown("### Notas de sesión (Gemini)")

    notes_file = st.file_uploader(
        "Archivo .docx de la carpeta Meet Recordings",
        type=["docx"], label_visibility="collapsed",
    )
    if notes_file and not st.session_state.notes_extracted:
        with st.spinner("Extrayendo contenido..."):
            try:
                result = extract_all(notes_file, notes_file.name)
                st.session_state.dev_points      = result["dev_points"]
                st.session_state.activities      = result["activities"]
                if result["fecha"]:
                    st.session_state.extracted_fecha = result["fecha"]
                st.session_state.notes_extracted = True
                st.rerun()
            except Exception as e:
                st.error(f"Error al leer el archivo: {e}")

    if st.session_state.notes_extracted:
        st.success(
            f"✓ {len(st.session_state.dev_points)} puntos · "
            f"{len(st.session_state.activities)} actividades extraídas"
        )

    st.divider()

    step_pill(7, "Genera el acta", done=st.session_state.docx_bytes is not None)
    st.markdown("### Generar acta")

    fecha_str     = fecha.strftime("%d-%m-%Y")
    alias         = empresa.split("(")[-1].replace(")", "").strip()[:12] if "(" in empresa else empresa[:12]
    acta_filename = f"Acta_{alias}_{fecha_str}.docx".replace(" ", "_")

    if st.button("📄 Generar acta .docx", type="primary", use_container_width=True):
        if not empresa or not jefe_cliente or not usuario_impl:
            st.error("Completa los campos obligatorios: Empresa, Jefe de proyecto y Usuario implementador.")
        else:
            with st.spinner("Generando acta..."):
                try:
                    header = {
                        "empresa":            empresa,
                        "plan":               plan,
                        "acta_num":           acta_num or fecha_str,
                        "fecha":              fecha_str,
                        "jefe_cliente":       jefe_cliente,
                        "email_jefe_cliente": email_jefe_cliente,
                        "tel_jefe_cliente":   tel_jefe_cliente,
                        "usuario_impl":       usuario_impl,
                        "email_impl":         email_impl,
                        "tel_usuario_impl":   tel_usuario_impl,
                        "jefe_rex":           jefe_rex,
                        "email_jefe_rex":     email_jefe_rex,
                        "tel_jefe_rex":       tel_jefe_rex,
                        "consultor":          consultor,
                        "email_consultor":    email_consultor,
                        "tel_consultor":      tel_consultor,
                    }
                    asistentes_list = [
                        {"nombre": str(row["nombre"]), "cargo": str(row.get("cargo", "")),
                         "gerencia": str(row.get("gerencia", ""))}
                        for _, row in edited.iterrows()
                        if str(row["nombre"]).strip()
                    ]

                    # Auto-guardar asistentes en el cliente ← CAMBIO 4
                    if client and asistentes_list:
                        updated = dict(client)
                        updated["asistentes"] = asistentes_list
                        save_client(updated)

                    st.session_state.docx_bytes   = generate_acta(
                        header=header,
                        asistentes=asistentes_list,
                        dev_points=st.session_state.dev_points,
                        activities=st.session_state.activities,
                    )
                    st.session_state.docx_filename = acta_filename
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")
                    st.exception(e)

    if st.session_state.docx_bytes:
        st.success("✓ Acta generada correctamente.")
        st.download_button(
            label="⬇️  Descargar acta .docx",
            data=st.session_state.docx_bytes,
            file_name=st.session_state.docx_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        st.divider()
        if st.button("🔁 Nueva acta", use_container_width=True):
            for k in ["dev_points", "activities", "notes_extracted",
                      "extracted_fecha", "docx_bytes", "docx_filename"]:
                v = st.session_state[k]
                st.session_state[k] = ([] if isinstance(v, list)
                                       else False if isinstance(v, bool)
                                       else "" if isinstance(v, str)
                                       else None)
            st.rerun()

aplicar_footer()
