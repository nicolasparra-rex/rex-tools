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

CLIENTES_PATH = Path(__file__).parent / "clientes.json"

st.set_page_config(
    page_title="Acta de Implementación · REX+",
    page_icon="📄",
    layout="wide",
)

st.markdown("""
<style>
    /* Fondo general */
    .stApp { background-color: #f5f6fa; }
    .block-container { padding-top: 1.5rem; }

    /* Títulos */
    h1 { font-size: 1.4rem !important; color: #0d1b3e !important; font-weight: 700 !important; }
    h2 { font-size: 1rem !important; color: #1a56db !important; font-weight: 600 !important;
         border-bottom: 2px solid #1a56db; padding-bottom: 4px; margin-top: 0.2rem !important; }

    /* Pill de paso */
    .step-pill {
        display: inline-flex; align-items: center; gap: 8px;
        background: white; border: 1px solid #dde3f0;
        border-radius: 99px; padding: 5px 14px 5px 6px;
        margin-bottom: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    .step-num {
        background: #1a56db; color: white; border-radius: 50%;
        width: 22px; height: 22px; display: inline-flex;
        align-items: center; justify-content: center;
        font-size: 11px; font-weight: 700; flex-shrink: 0;
    }
    .step-num.done { background: #16a34a; }
    .step-label { font-size: 12px; font-weight: 600; color: #0d1b3e; }

    /* Labels */
    label { color: #374151 !important; font-size: 0.82rem !important; font-weight: 500 !important; }

    /* Inputs focus */
    input:focus, textarea:focus {
        border-color: #1a56db !important;
        box-shadow: 0 0 0 2px rgba(26,86,219,0.12) !important;
    }

    /* Botón primario */
    div[data-testid="stButton"] > button[kind="primary"] {
        background: #1a56db !important; color: white !important;
        border: none !important; border-radius: 8px !important;
        font-weight: 600 !important;
    }
    div[data-testid="stButton"] > button[kind="primary"]:hover { background: #1447c0 !important; }

    /* Botón secundario */
    div[data-testid="stButton"] > button {
        background: white !important; color: #1a56db !important;
        border: 1.5px solid #1a56db !important; border-radius: 8px !important;
        font-weight: 500 !important;
    }
    div[data-testid="stButton"] > button:hover { background: #eef2fd !important; }

    /* Botón descarga */
    div[data-testid="stDownloadButton"] > button {
        background: #16a34a !important; color: white !important;
        border: none !important; border-radius: 8px !important;
        font-weight: 600 !important;
    }
    div[data-testid="stDownloadButton"] > button:hover { background: #15803d !important; }

    /* Selectbox */
    div[data-testid="stSelectbox"] > div > div {
        border-radius: 6px !important; background: white !important;
    }

    /* Alertas */
    div[data-testid="stAlert"] { border-radius: 8px !important; }

    hr { border-color: #e2e8f0 !important; }
    div[data-testid="stDataEditor"] { border-radius: 8px !important; }
</style>
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
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── Header ────────────────────────────────────────────────────────────────────

st.title("📄 Acta de Implementación · REX+")
st.divider()

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

    col_sel, col_del = st.columns([4, 1])
    with col_sel:
        selected = st.selectbox(
            "Cliente guardado",
            options, index=0, label_visibility="collapsed",
        )
    with col_del:
        if selected != "— Nuevo cliente —":
            if st.button("🗑", use_container_width=True, help="Eliminar cliente"):
                delete_client(next(c["empresa"] for c in client_list if c["nombre_display"] == selected))
                st.rerun()

    client = next((c for c in client_list if c["nombre_display"] == selected), None)

    st.divider()

    # ── PASO 2: Datos del cliente ─────────────────────────────────────────────
    step_pill(2, "Datos del cliente")
    st.markdown("## Datos del cliente")

    empresa = st.text_input(
        "Empresa *",
        value=client["empresa"] if client else "",
        placeholder="RE-XXXX - Nombre empresa (ALIAS)",
    )
    jefe_cliente = st.text_input(
        "Jefe de proyecto cliente *",
        value=client["jefe_cliente"] if client else "",
        placeholder="Nombre completo en mayúsculas",
    )
    col1, col2 = st.columns(2)
    with col1:
        usuario_impl = st.text_input(
            "Usuario implementador *",
            value=client["usuario_impl"] if client else "",
        )
    with col2:
        email_impl = st.text_input(
            "Email implementador",
            value=client["email_impl"] if client else "",
            placeholder="nombre@empresa.cl",
        )

    st.divider()

    # ── PASO 3: Equipo REX+ ───────────────────────────────────────────────────
    step_pill(3, "Equipo REX+")
    st.markdown("## Equipo REX+")

    col1, col2 = st.columns(2)
    with col1:
        jefe_rex = st.text_input(
            "Jefe de proyecto REX+",
            value=client["jefe_rex"] if client else "GABRIELA AHUMADA",
        )
        email_jefe_rex = st.text_input(
            "Email jefe REX+",
            value=client["email_jefe_rex"] if client else "Gabriela.ahumada@visma.com",
        )
    with col2:
        consultor = st.text_input(
            "Consultor REX",
            value=client["consultor"] if client else "DIEGO GALVEZ",
        )
        horas = st.number_input(
            "Horas de sesión", min_value=1, max_value=8,
            value=int(client["horas"]) if client else 4,
        )

    st.divider()

    # ── PASO 4: Sesión ────────────────────────────────────────────────────────
    step_pill(4, "Datos de la sesión")
    st.markdown("## Datos de la sesión")

    plan_opts = ["BASE", "FULL", "PREMIUM"]
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

    # ── Guardar cliente ───────────────────────────────────────────────────────
    with st.expander("💾 Guardar datos de este cliente"):
        nombre_display = st.text_input(
            "Nombre corto del cliente",
            value=client["nombre_display"] if client else "",
            placeholder="Ej: CYGNUS",
        )
        if st.button("Guardar cliente", use_container_width=True):
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
                st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# COLUMNA DERECHA
# ═══════════════════════════════════════════════════════════════════════════════
with col_right:

    # ── PASO 6: Notas de Gemini ───────────────────────────────────────────────
    step_pill(6, "Sube las notas de Gemini", done=st.session_state.notes_extracted)
    st.markdown("## Notas de sesión (Gemini)")

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
