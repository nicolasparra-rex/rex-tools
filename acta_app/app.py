"""
app.py  —  Acta de Implementación REX+
Ejecutar con: python -m streamlit run app.py
"""

import json
import datetime
from pathlib import Path

import streamlit as st
import pandas as pd

from extractor import extract_all
from generator import generate_acta

# Resolver ruta correctamente tanto en ejecución directa como desde pages/
_BASE_DIR = Path(__file__).parent
# Si se ejecuta desde pages/, subir un nivel y entrar a acta_app/
if _BASE_DIR.name == "pages":
    _BASE_DIR = _BASE_DIR.parent / "acta_app"

CLIENTES_PATH = _BASE_DIR / "clientes.json"
EQUIPO_PATH   = _BASE_DIR / "equipo.json"

st.set_page_config(
    page_title="Acta de Implementación · REX+",
    page_icon="📄",
    layout="wide",
)

# ── Paleta Rex+ Tools ────────────────────────────────────────────────────────
REX_NAVY  = "#1a2744"
REX_CYAN  = "#00c4cc"
REX_LIGHT = "#f0fafb"

st.markdown(f"""
<style>
    /* Fondo general */
    .stApp {{ background-color: #f8fafc; }}
    .block-container {{ padding-top: 0 !important; max-width: 100% !important; }}

    /* Ocultar header nativo de Streamlit */
    header[data-testid="stHeader"] {{ display: none !important; }}

    /* Barra superior personalizada */
    .rex-header {{
        background: linear-gradient(90deg, {REX_NAVY} 0%, #1e3460 100%);
        padding: 0 2rem;
        height: 56px;
        display: flex;
        align-items: center;
        gap: 14px;
        margin-bottom: 1.5rem;
        position: sticky;
        top: 0;
        z-index: 999;
        box-shadow: 0 2px 8px rgba(0,0,0,0.2);
    }}
    .rex-logo {{
        background: white;
        color: {REX_NAVY};
        font-weight: 800;
        font-size: 13px;
        padding: 5px 8px;
        border-radius: 6px;
        letter-spacing: 0.5px;
    }}
    .rex-divider {{
        width: 1px; height: 28px;
        background: rgba(255,255,255,0.25);
    }}
    .rex-title {{
        color: white;
        font-size: 15px;
        font-weight: 600;
        letter-spacing: 0.3px;
        flex: 1;
    }}
    .rex-badge {{
        background: {REX_CYAN};
        color: {REX_NAVY};
        font-size: 10px;
        font-weight: 800;
        padding: 3px 10px;
        border-radius: 99px;
        letter-spacing: 1px;
        text-transform: uppercase;
    }}

    /* Sidebar */
    section[data-testid="stSidebar"] {{
        background-color: {REX_NAVY} !important;
    }}
    section[data-testid="stSidebar"] * {{
        color: rgba(255,255,255,0.75) !important;
    }}
    section[data-testid="stSidebar"] .st-emotion-cache-1cypcdb {{
        background-color: {REX_CYAN} !important;
        border-radius: 6px !important;
    }}

    /* Títulos */
    h1 {{ display: none !important; }}
    h2 {{
        font-size: 1.05rem !important; color: {REX_NAVY} !important;
        font-weight: 700 !important;
        border-bottom: 2px solid {REX_CYAN};
        padding-bottom: 5px; margin-top: 0.5rem !important;
    }}

    /* Pills de paso */
    .step-pill {{
        display: inline-flex; align-items: center; gap: 8px;
        background: white; border: 1px solid #dde3f0;
        border-radius: 99px; padding: 5px 14px 5px 6px;
        margin-bottom: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }}
    .step-num {{
        background: {REX_NAVY}; color: white; border-radius: 50%;
        width: 22px; height: 22px; display: inline-flex;
        align-items: center; justify-content: center;
        font-size: 11px; font-weight: 700; flex-shrink: 0;
    }}
    .step-num.done {{ background: {REX_CYAN}; color: {REX_NAVY}; }}
    .step-label {{ font-size: 12px; font-weight: 600; color: {REX_NAVY}; }}

    /* Labels */
    label {{ color: #374151 !important; font-size: 0.82rem !important; font-weight: 500 !important; }}

    /* Inputs focus */
    input:focus, textarea:focus {{
        border-color: {REX_CYAN} !important;
        box-shadow: 0 0 0 2px rgba(0,196,204,0.2) !important;
    }}

    /* Botón primario */
    div[data-testid="stButton"] > button[kind="primary"] {{
        background: {REX_NAVY} !important; color: white !important;
        border: none !important; border-radius: 20px !important;
        font-weight: 600 !important;
    }}
    div[data-testid="stButton"] > button[kind="primary"]:hover {{
        background: {REX_CYAN} !important; color: {REX_NAVY} !important;
    }}

    /* Botón secundario */
    div[data-testid="stButton"] > button {{
        background: white !important; color: {REX_NAVY} !important;
        border: 1.5px solid {REX_NAVY} !important; border-radius: 20px !important;
        font-weight: 500 !important;
    }}
    div[data-testid="stButton"] > button:hover {{
        background: {REX_LIGHT} !important; border-color: {REX_CYAN} !important;
        color: {REX_NAVY} !important;
    }}

    /* Botón descarga */
    div[data-testid="stDownloadButton"] > button {{
        background: {REX_CYAN} !important; color: {REX_NAVY} !important;
        border: none !important; border-radius: 20px !important;
        font-weight: 700 !important;
    }}
    div[data-testid="stDownloadButton"] > button:hover {{
        opacity: 0.85 !important;
    }}

    /* Selectbox */
    div[data-testid="stSelectbox"] > div > div {{
        border-radius: 6px !important; background: white !important;
    }}

    /* Alertas */
    div[data-testid="stAlert"] {{ border-radius: 8px !important; }}

    /* File uploader — borde punteado cyan */
    div[data-testid="stFileUploader"] {{
        border: 2px dashed {REX_CYAN} !important;
        border-radius: 10px !important;
        background: {REX_LIGHT} !important;
        padding: 8px !important;
    }}

    hr {{ border-color: #e2e8f0 !important; }}
    div[data-testid="stDataEditor"] {{ border-radius: 8px !important; }}
</style>
""", unsafe_allow_html=True)

# Barra superior
st.markdown(f"""
<div class="rex-header">
    <span class="rex-logo">Rex+</span>
    <span class="rex-divider"></span>
    <span class="rex-title">Acta de Implementación</span>
    <span class="rex-badge">IMPLEMENTACIÓN</span>
</div>
""", unsafe_allow_html=True)

# ── Helpers ──────────────────────────────────────────────────────────────────

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

def add_equipo_member(rol: str, nombre: str, email: str):
    data = load_equipo()
    lista = data[rol]
    if not any(m["nombre"].lower() == nombre.lower() for m in lista):
        lista.append({"nombre": nombre, "email": email})
        save_equipo(data)

def delete_equipo_member(rol: str, nombre: str):
    data = load_equipo()
    data[rol] = [m for m in data[rol] if m["nombre"] != nombre]
    save_equipo(data)

def step_pill(num: int, label: str, done: bool = False):
    cls = "done" if done else ""
    icon = "✓" if done else str(num)
    st.markdown(f"""
    <div class="step-pill">
        <span class="step-num {cls}">{icon}</span>
        <span class="step-label">{label}</span>
    </div>""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────

for k, v in {
    "dev_points": [], "activities": [],
    "notes_extracted": False, "extracted_fecha": "",
    "docx_bytes": None, "docx_filename": "",
    "edit_mode": False, "last_selected": "",
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── Header ya renderizado via HTML ───────────────────────────────────────────

# ── Layout: columna izquierda (formulario) | columna derecha (notas + generar)
col_form, col_right = st.columns([3, 2], gap="large")

# ═══════════════════════════════════════════════════════════════════════════════
# COLUMNA IZQUIERDA
# ═══════════════════════════════════════════════════════════════════════════════
with col_form:

    # ── PASO 1: Cliente ───────────────────────────────────────────────────────
    has_client = bool(st.session_state.get("_empresa", ""))
    step_pill(1, "Selecciona o ingresa el cliente", done=has_client)

    data = load_clients()
    client_list = data["clientes"]
    client_names = [c["nombre_display"] for c in client_list]
    options = ["— Nuevo cliente —"] + client_names

    col_sel, col_edit, col_del = st.columns([4, 1, 1])
    with col_sel:
        selected = st.selectbox(
            "Cliente guardado",
            options, index=0, label_visibility="collapsed",
        )
    # Resetear edit_mode si cambió el cliente seleccionado
    if selected != st.session_state.last_selected:
        st.session_state.edit_mode = False
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

    client = next((c for c in client_list if c["nombre_display"] == selected), None)

    # Modo lectura vs edición
    is_new    = selected == "— Nuevo cliente —"
    editable  = is_new or st.session_state.edit_mode

    if client and not editable:
        # ── Vista solo lectura ────────────────────────────────────────────────
        st.info("🔒 Datos del cliente — haz clic en **🔓 Editar** para modificar")

    st.divider()

    # ── PASO 2: Datos del cliente ─────────────────────────────────────────────
    step_pill(2, "Datos del cliente")
    st.markdown("## Datos del cliente")

    empresa = st.text_input(
        "Empresa *",
        value=client["empresa"] if client else "",
        placeholder="RE-XXXX - Nombre empresa (ALIAS)",
        disabled=not editable,
    )
    jefe_cliente = st.text_input(
        "Jefe de proyecto cliente *",
        value=client["jefe_cliente"] if client else "",
        placeholder="Nombre completo en mayúsculas",
        disabled=not editable,
    )
    col1, col2 = st.columns(2)
    with col1:
        usuario_impl = st.text_input(
            "Usuario implementador *",
            value=client["usuario_impl"] if client else "",
            disabled=not editable,
        )
    with col2:
        email_impl = st.text_input(
            "Email implementador",
            value=client["email_impl"] if client else "",
            placeholder="nombre@empresa.cl",
            disabled=not editable,
        )

    st.divider()

    # ── PASO 3: Equipo REX+ ───────────────────────────────────────────────────
    step_pill(3, "Equipo REX+")
    st.markdown("## Equipo REX+")

    equipo = load_equipo()
    jefes_list     = equipo.get("jefes", [])
    consultores_list = equipo.get("consultores", [])

    jefes_nombres     = [m["nombre"] for m in jefes_list]
    consultores_nombres = [m["nombre"] for m in consultores_list]

    col1, col2 = st.columns(2)
    with col1:
        jefe_opts = ["— Nuevo —"] + jefes_nombres
        jefe_default = client["jefe_rex"] if client else (jefes_nombres[0] if jefes_nombres else "— Nuevo —")
        jefe_idx = jefe_opts.index(jefe_default) if jefe_default in jefe_opts else 0
        jefe_sel = st.selectbox("Jefe de proyecto REX+", jefe_opts, index=jefe_idx)
        if jefe_sel == "— Nuevo —":
            jefe_rex = st.text_input("Nombre jefe *", placeholder="NOMBRE APELLIDO")
            email_jefe_rex = st.text_input("Email jefe", placeholder="nombre@visma.com")
        else:
            jefe_data = next(m for m in jefes_list if m["nombre"] == jefe_sel)
            jefe_rex = jefe_sel
            email_jefe_rex = st.text_input("Email jefe", value=jefe_data["email"])

    with col2:
        cons_opts = ["— Nuevo —"] + consultores_nombres
        cons_default = client["consultor"] if client else (consultores_nombres[0] if consultores_nombres else "— Nuevo —")
        cons_idx = cons_opts.index(cons_default) if cons_default in cons_opts else 0
        cons_sel = st.selectbox("Consultor REX", cons_opts, index=cons_idx)
        if cons_sel == "— Nuevo —":
            consultor = st.text_input("Nombre consultor *", placeholder="NOMBRE APELLIDO")
            email_consultor = st.text_input("Email consultor", placeholder="nombre@visma.com")
        else:
            cons_data = next(m for m in consultores_list if m["nombre"] == cons_sel)
            consultor = cons_sel
            email_consultor = st.text_input("Email consultor", value=cons_data["email"], key="email_cons")

        horas = st.number_input(
            "Horas de sesión", min_value=1, max_value=8,
            value=int(client["horas"]) if client else 4,
        )

    # Guardar nuevos miembros del equipo
    col_gj, col_gc = st.columns(2)
    with col_gj:
        if jefe_sel == "— Nuevo —" and jefe_rex:
            if st.button("💾 Guardar jefe", use_container_width=True):
                add_equipo_member("jefes", jefe_rex, email_jefe_rex)
                st.success(f"✓ Jefe '{jefe_rex}' guardado.")
                st.rerun()
    with col_gc:
        if cons_sel == "— Nuevo —" and consultor:
            if st.button("💾 Guardar consultor", use_container_width=True):
                add_equipo_member("consultores", consultor, email_consultor)
                st.success(f"✓ Consultor '{consultor}' guardado.")
                st.rerun()

    # Gestión del equipo guardado
    with st.expander("🗑 Gestionar equipo guardado"):
        col_ej, col_ec = st.columns(2)
        with col_ej:
            st.markdown("**Jefes**")
            for m in jefes_list:
                c1, c2 = st.columns([3,1])
                c1.write(m["nombre"])
                if c2.button("🗑", key=f"del_j_{m['nombre']}"):
                    delete_equipo_member("jefes", m["nombre"])
                    st.rerun()
        with col_ec:
            st.markdown("**Consultores**")
            for m in consultores_list:
                c1, c2 = st.columns([3,1])
                c1.write(m["nombre"])
                if c2.button("🗑", key=f"del_c_{m['nombre']}"):
                    delete_equipo_member("consultores", m["nombre"])
                    st.rerun()

    st.divider()

    # ── PASO 4: Sesión ────────────────────────────────────────────────────────
    step_pill(4, "Datos de la sesión")
    st.markdown("## Datos de la sesión")

    plan_opts = ["Full", "Estándar", "Express", "Base", "Casino con instalación", "Casino sin instalación", "Asistencia solo Marcaje app"]
    plan_idx = plan_opts.index(client["plan"]) if client and client.get("plan") in plan_opts else 0
    col1, col2, col3 = st.columns(3)
    with col1:
        plan = st.selectbox("Plan", plan_opts, index=plan_idx)
    with col2:
        acta_num = st.text_input(
            "N° de acta",
            value=st.session_state.extracted_fecha,
            placeholder="Ej: 27-04-2026",
        )
    with col3:
        fecha = st.date_input("Fecha de sesión", value=datetime.date.today())

    st.divider()

    # ── PASO 5: Asistentes ────────────────────────────────────────────────────
    step_pill(5, "Asistentes a la sesión")
    st.markdown("## Asistentes a la sesión")

    default_asistentes = [
        {"nombre": client["usuario_impl"] if client else "", "cargo": "", "gerencia": ""},
        {"nombre": client["jefe_cliente"] if client else "",  "cargo": "", "gerencia": ""},
        {"nombre": "", "cargo": "", "gerencia": ""},
        {"nombre": "", "cargo": "", "gerencia": ""},
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

    # ── Guardar / actualizar cliente ──────────────────────────────────────────
    st.divider()
    if is_new:
        with st.expander("💾 Guardar nuevo cliente"):
            nombre_display = st.text_input(
                "Nombre corto del cliente",
                value="",
                placeholder="Ej: CYGNUS",
            )
            if st.button("Guardar cliente", use_container_width=True, type="primary"):
                if not empresa:
                    st.error("Ingresa el nombre de la empresa.")
                else:
                    save_client({
                        "nombre_display": nombre_display or empresa[:20],
                        "empresa": empresa, "plan": plan,
                        "jefe_cliente": jefe_cliente,
                        "usuario_impl": usuario_impl, "email_impl": email_impl,
                        "jefe_rex": jefe_rex, "email_jefe_rex": email_jefe_rex,
                        "consultor": consultor, "horas": int(horas),
                        "keyword_drive": nombre_display.lower() if nombre_display else "",
                    })
                    st.success(f"✓ Cliente '{nombre_display}' guardado.")
                    st.session_state.edit_mode = False
                    st.rerun()
    elif st.session_state.edit_mode:
        if st.button("💾 Guardar cambios", use_container_width=True, type="primary"):
            if not empresa:
                st.error("Ingresa el nombre de la empresa.")
            else:
                save_client({
                    "nombre_display": client["nombre_display"],
                    "empresa": empresa, "plan": plan,
                    "jefe_cliente": jefe_cliente,
                    "usuario_impl": usuario_impl, "email_impl": email_impl,
                    "jefe_rex": jefe_rex, "email_jefe_rex": email_jefe_rex,
                    "consultor": consultor, "horas": int(horas),
                    "keyword_drive": client.get("keyword_drive", ""),
                })
                st.success(f"✓ Cambios guardados para **{client['nombre_display']}**.")
                st.session_state.edit_mode = False
                st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# COLUMNA DERECHA
# ═══════════════════════════════════════════════════════════════════════════════
with col_right:

    # ── PASO 6: Notas de Gemini ───────────────────────────────────────────────
    step_pill(6, "Sube las notas de Gemini", done=st.session_state.notes_extracted)
    st.markdown("## Notas de sesión (Gemini)")

    import sys as _sys
    _sys.path.insert(0, str(Path(__file__).parent))
    try:
        from drive_client import is_configured, search_gemini_notes, download_docx
    except ModuleNotFoundError:
        def is_configured(): return False
        def search_gemini_notes(k): return []
        def download_docx(i): return None

    # ── Modo: Drive o upload manual ──────────────────────────────────────────
    drive_ok = is_configured()
    modo = st.radio(
        "Fuente",
        ["📁 Buscar en Google Drive", "💻 Subir archivo manualmente"],
        horizontal=True,
        label_visibility="collapsed",
        index=0 if drive_ok else 1,
    )

    if modo == "📁 Buscar en Google Drive":
        if not drive_ok:
            st.warning("⚠️ Drive no configurado. Ve a la sección **Configurar Drive** al final de la app.")
        else:
            keyword = client["keyword_drive"] if client else ""
            if not keyword and empresa:
                # Extraer keyword del nombre de empresa
                keyword = empresa.split("(")[-1].replace(")","").strip().lower() if "(" in empresa else empresa.split("-")[-1].strip().lower()

            col_kw, col_btn = st.columns([3,1])
            with col_kw:
                keyword = st.text_input("Buscar sesiones de", value=keyword, placeholder="ej: cygnus")
            with col_btn:
                st.write("")
                st.write("")
                buscar = st.button("🔍 Buscar", use_container_width=True)

            if buscar and keyword:
                with st.spinner(f"Buscando sesiones de '{keyword}' en Drive..."):
                    try:
                        st.session_state["drive_files"] = search_gemini_notes(keyword)
                    except Exception as e:
                        st.error(f"Error al buscar en Drive: {e}")

            files = st.session_state.get("drive_files", [])
            if files:
                opciones = {f"{f['fecha_str']} — {f['name'][:50]}": f for f in files}
                seleccion = st.selectbox("Selecciona la sesión", list(opciones.keys()), label_visibility="collapsed")
                archivo_sel = opciones[seleccion]

                if st.button("⬇️ Cargar sesión", type="primary", use_container_width=True):
                    with st.spinner("Descargando desde Drive..."):
                        try:
                            buffer = download_docx(archivo_sel["id"])
                            result = extract_all(buffer, archivo_sel["name"])
                            st.session_state.dev_points = result["dev_points"]
                            st.session_state.activities = result["activities"]
                            if result["fecha"]:
                                st.session_state.extracted_fecha = result["fecha"]
                            st.session_state.notes_extracted = True
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error al cargar el archivo: {e}")
            elif st.session_state.get("drive_files") is not None:
                st.info("No se encontraron sesiones. Intenta con otro nombre.")

    else:
        notes_file = st.file_uploader(
            "Archivo .docx de la carpeta Meet Recordings",
            type=["docx"], label_visibility="collapsed",
        )
        if notes_file and not st.session_state.notes_extracted:
            with st.spinner("Extrayendo contenido..."):
                try:
                    result = extract_all(notes_file, notes_file.name)
                    st.session_state.dev_points = result["dev_points"]
                    st.session_state.activities = result["activities"]
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

    # ── PASO 7: Generar acta ──────────────────────────────────────────────────
    step_pill(7, "Genera el acta", done=st.session_state.docx_bytes is not None)
    st.markdown("## Generar acta")

    fecha_str = fecha.strftime("%d-%m-%Y")
    alias = empresa.split("(")[-1].replace(")","").strip()[:12] if "(" in empresa else empresa[:12]
    acta_filename = f"Acta_{alias}_{fecha_str}.docx".replace(" ", "_")

    if st.button("📄 Generar acta .docx", type="primary", use_container_width=True):
        if not empresa or not jefe_cliente or not usuario_impl:
            st.error("Completa los campos obligatorios: Empresa, Jefe de proyecto y Usuario implementador.")
        else:
            with st.spinner("Generando acta..."):
                try:
                    header = {
                        "empresa": empresa, "plan": plan,
                        "acta_num": acta_num or fecha_str,
                        "fecha": fecha_str,
                        "jefe_cliente": jefe_cliente,
                        "usuario_impl": usuario_impl, "email_impl": email_impl,
                        "jefe_rex": jefe_rex, "email_jefe_rex": email_jefe_rex,
                        "consultor": consultor,
                    }
                    asistentes_list = [
                        {"nombre": row["nombre"], "cargo": row.get("cargo",""), "gerencia": row.get("gerencia","")}
                        for _, row in edited.iterrows()
                        if str(row["nombre"]).strip()
                    ]
                    st.session_state.docx_bytes = generate_acta(
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
            for k in ["dev_points","activities","notes_extracted",
                      "extracted_fecha","docx_bytes","docx_filename"]:
                st.session_state[k] = [] if isinstance(st.session_state[k], list) else (False if isinstance(st.session_state[k], bool) else ("" if isinstance(st.session_state[k], str) else None))
            st.rerun()

st.divider()

# ── Configurar Drive ──────────────────────────────────────────────────────────
with st.expander("⚙️ Configurar acceso a Google Drive"):
    st.markdown("""
    ### Conectar Google Drive (una sola vez)

    **Paso 1** — Ve a [Google OAuth Playground](https://developers.google.com/oauthplayground/)

    **Paso 2** — En la lista de APIs, busca **"Drive API v3"** y selecciona:
    - `https://www.googleapis.com/auth/drive.readonly`

    **Paso 3** — Haz clic en **"Authorize APIs"** → inicia sesión con tu cuenta Google

    **Paso 4** — Haz clic en **"Exchange authorization code for tokens"**

    **Paso 5** — Copia el valor de **Refresh token**

    **Paso 6** — Pégalo aquí:
    """)

    refresh_input = st.text_input(
        "Refresh token de Google",
        type="password",
        placeholder="1//04xxx...",
    )
    if st.button("💾 Guardar token", use_container_width=True):
        if refresh_input:
            env_path = Path(__file__).parent / ".env"
            with open(env_path, "w") as f:
                f.write(f"GOOGLE_REFRESH_TOKEN={refresh_input}\n")
            st.success("✓ Token guardado. Reinicia la app para activar Drive.")
        else:
            st.error("Pega el refresh token antes de guardar.")
