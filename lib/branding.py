"""
Módulo compartido Rex+ - Branding, CSS y utilidades comunes.
Se importa desde app.py y desde cada página en pages/.
"""

import base64
from pathlib import Path
import streamlit as st


# ─────────────────────────────────────────────
#  COLORES REX+
# ─────────────────────────────────────────────
REX_AZUL        = "#1EBBEF"
REX_AZUL_OSCURO = "#1A3A5F"
REX_LIMA        = "#C5E86C"
REX_NARANJA     = "#F5A623"
REX_ROJO        = "#D8594B"
REX_GRIS_CLARO  = "#F5F9FC"
REX_GRIS_MEDIO  = "#8B9DAE"


# ─────────────────────────────────────────────
#  CSS GLOBAL
# ─────────────────────────────────────────────
CSS_PERSONALIZADO = """
<style>
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    .main .block-container {
        padding-top: 2rem !important;
        padding-left: 2rem;
        padding-right: 2rem;
        max-width: 1400px;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: #1A3A5F;
    }
    [data-testid="stSidebar"] .sidebar-logo {
        background: white;
        padding: 12px 16px;
        border-radius: 10px;
        margin: 0.5rem 0 1.5rem 0;
        text-align: center;
    }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] .stRadio label {
        color: white !important;
    }
    [data-testid="stSidebar"] [data-testid="stSidebarNav"] a {
        color: rgba(255,255,255,0.85) !important;
        padding: 10px 16px !important;
        margin: 4px 8px !important;
        border-radius: 8px !important;
        transition: all 0.2s ease;
    }
    [data-testid="stSidebar"] [data-testid="stSidebarNav"] a:hover {
        background: rgba(30, 187, 239, 0.15) !important;
        color: #1EBBEF !important;
    }
    [data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"] {
        background: #1EBBEF !important;
        color: white !important;
        font-weight: 600 !important;
    }

    /* Header interno */
    .rex-header {
        background: linear-gradient(135deg, #1A3A5F 0%, #2E5A8C 100%);
        padding: 1.25rem 2rem;
        margin-bottom: 2rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(26, 58, 95, 0.15);
    }
    .rex-header-left {
        display: flex;
        align-items: center;
        gap: 1rem;
    }
    .rex-logo {
        height: 42px;
        background: white;
        padding: 6px 14px;
        border-radius: 8px;
    }
    .rex-divider {
        width: 1px;
        height: 28px;
        background: rgba(255,255,255,0.25);
    }
    .rex-header-title {
        color: white;
        font-size: 1.05rem;
        font-weight: 500;
    }
    .rex-header-badge {
        background: #1EBBEF;
        color: white;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 600;
    }

    /* Hero */
    .rex-hero {
        margin: 0 0 2rem 0;
    }
    .rex-hero h1 {
        color: #1A3A5F;
        font-size: 1.75rem;
        font-weight: 700;
        margin: 0 0 0.25rem 0;
    }
    .rex-hero p {
        color: #8B9DAE;
        font-size: 0.95rem;
        margin: 0;
    }

    /* Cards del home */
    .rex-tool-card {
        background: white;
        border: 1px solid #E8EEF3;
        border-radius: 16px;
        padding: 2rem;
        margin-bottom: 1.5rem;
        transition: all 0.25s ease;
        box-shadow: 0 2px 4px rgba(26, 58, 95, 0.04);
        display: block;
        text-decoration: none !important;
        color: inherit !important;
    }
    .rex-tool-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 24px rgba(30, 187, 239, 0.15);
        border-color: #1EBBEF;
    }
    .rex-tool-icon {
        width: 54px;
        height: 54px;
        background: linear-gradient(135deg, #1EBBEF 0%, #0095C9 100%);
        border-radius: 14px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.75rem;
        margin-bottom: 1rem;
        color: white;
    }
    .rex-tool-icon.lime {
        background: linear-gradient(135deg, #C5E86C 0%, #9DC93E 100%);
    }
    .rex-tool-icon.dark {
        background: linear-gradient(135deg, #1A3A5F 0%, #2E5A8C 100%);
    }
    .rex-tool-icon.orange {
        background: linear-gradient(135deg, #F5A623 0%, #D68910 100%);
    }
    .rex-tool-title {
        color: #1A3A5F;
        font-size: 1.15rem;
        font-weight: 700;
        margin: 0 0 0.5rem 0;
    }
    .rex-tool-desc {
        color: #6B7C8E;
        font-size: 0.9rem;
        line-height: 1.5;
        margin: 0 0 1rem 0;
    }
    .rex-tool-cta {
        color: #1EBBEF;
        font-weight: 600;
        font-size: 0.85rem;
        display: inline-flex;
        align-items: center;
        gap: 4px;
    }
    .rex-tool-tag {
        display: inline-block;
        background: #F5F9FC;
        color: #8B9DAE;
        padding: 3px 10px;
        border-radius: 12px;
        font-size: 0.7rem;
        font-weight: 600;
        margin-right: 6px;
        letter-spacing: 0.3px;
    }
    .rex-tool-tag.activa {
        background: #EEF7FD;
        color: #1EBBEF;
    }
    .rex-tool-tag.proximo {
        background: #FEF7EC;
        color: #F5A623;
    }

    /* Métricas */
    [data-testid="stMetric"] {
        background: white;
        border: 1px solid #E8EEF3;
        border-radius: 12px;
        padding: 1.25rem;
        box-shadow: 0 1px 3px rgba(26, 58, 95, 0.04);
        transition: all 0.2s ease;
    }
    [data-testid="stMetric"]:hover {
        box-shadow: 0 4px 12px rgba(30, 187, 239, 0.12);
        border-color: #1EBBEF;
    }
    [data-testid="stMetricLabel"] {
        color: #8B9DAE !important;
        font-size: 0.8rem !important;
        font-weight: 500 !important;
    }
    [data-testid="stMetricValue"] {
        color: #1A3A5F !important;
        font-size: 1.75rem !important;
        font-weight: 700 !important;
    }

    /* File uploader */
    [data-testid="stFileUploader"] section {
        background: #F5F9FC;
        border: 2px dashed #1EBBEF !important;
        border-radius: 12px;
        padding: 2rem;
    }
    [data-testid="stFileUploader"] section:hover {
        background: #EEF7FD;
        border-color: #1A3A5F !important;
    }
    [data-testid="stFileUploader"] button {
        background: #1A3A5F !important;
        color: white !important;
        border: none !important;
        border-radius: 20px !important;
        font-weight: 600 !important;
    }
    [data-testid="stFileUploader"] button:hover {
        background: #1EBBEF !important;
    }

    /* Botones */
    .stButton > button, .stDownloadButton > button {
        background: #1A3A5F;
        color: white;
        border: none;
        border-radius: 20px;
        padding: 0.5rem 1.5rem;
        font-weight: 600;
        font-size: 0.9rem;
        transition: all 0.2s ease;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: #1EBBEF;
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(30, 187, 239, 0.25);
    }
    .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
        background: #1EBBEF;
    }
    .stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {
        background: #1A3A5F;
    }

    /* Alertas */
    .stAlert {
        border-radius: 10px;
        border-left-width: 4px;
    }

    /* Expanders */
    [data-testid="stExpander"] {
        background: white;
        border: 1px solid #E8EEF3;
        border-radius: 10px;
        margin-bottom: 1rem;
    }

    /* DataFrame */
    [data-testid="stDataFrame"] {
        border: 1px solid #E8EEF3;
        border-radius: 10px;
        overflow: hidden;
    }

    /* Footer */
    .rex-footer {
        margin-top: 3rem;
        padding: 1.5rem 0;
        border-top: 1px solid #E8EEF3;
        text-align: center;
        color: #8B9DAE;
        font-size: 0.8rem;
    }
    .rex-footer strong {
        color: #1EBBEF;
    }
</style>
"""


def _cargar_logo_base64():
    """Carga el logo como base64 para incrustarlo en el HTML."""
    # Buscar el logo en varias rutas posibles (según desde dónde se importe)
    posibles = [
        Path(__file__).parent.parent / "assets" / "logo.png",
        Path(__file__).parent / "assets" / "logo.png",
        Path("assets") / "logo.png",
    ]
    for ruta in posibles:
        if ruta.exists():
            with open(ruta, "rb") as f:
                return base64.b64encode(f.read()).decode()
    return None


def aplicar_branding(titulo_pagina: str = "", badge: str = "PRODUCCION"):
    """Aplica el CSS y el header Rex+ a la página actual.

    Llamar al inicio de cada página (app.py o pages/*.py).
    """
    st.markdown(CSS_PERSONALIZADO, unsafe_allow_html=True)

    logo_b64 = _cargar_logo_base64()
    if logo_b64:
        logo_html = f'<img src="data:image/png;base64,{logo_b64}" class="rex-logo" alt="Rex+"/>'
    else:
        logo_html = '<div class="rex-logo" style="color:#1EBBEF;font-weight:800;font-size:1.4rem;padding:8px 14px;">Rex+</div>'

    titulo_header = titulo_pagina if titulo_pagina else "Rex+ Tools"

    header_html = (
        '<div class="rex-header">'
        '<div class="rex-header-left">'
        + logo_html
        + '<div class="rex-divider"></div>'
        f'<div class="rex-header-title">{titulo_header}</div>'
        '</div>'
        f'<div class="rex-header-badge">{badge}</div>'
        '</div>'
    )
    st.markdown(header_html, unsafe_allow_html=True)


def aplicar_footer():
    """Agrega el footer Rex+ al final de la página."""
    st.markdown(
        '<div class="rex-footer">Powered by <strong>Rex+</strong> · Rex+ Tools · v1.0</div>',
        unsafe_allow_html=True,
    )


def hero(titulo: str, descripcion: str = "", icono: str = ""):
    """Muestra un encabezado hero con título y descripción."""
    titulo_mostrar = f"{icono} {titulo}" if icono else titulo
    st.markdown(
        f'<div class="rex-hero">'
        f'<h1>{titulo_mostrar}</h1>'
        f'<p>{descripcion}</p>'
        '</div>',
        unsafe_allow_html=True,
    )
