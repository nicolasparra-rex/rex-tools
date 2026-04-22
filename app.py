"""
Rex+ Tools - Página principal (Landing / Dashboard)
Muestra un dashboard con cards clickeables hacia cada herramienta.
"""

import streamlit as st
from lib.branding import aplicar_branding, aplicar_footer, hero

st.set_page_config(
    page_title="Rex+ Tools",
    page_icon="🛠️",
    layout="wide",
    initial_sidebar_state="expanded",
)

aplicar_branding(titulo_pagina="Rex+ Tools", badge="PRODUCCION")

hero(
    titulo="Bienvenido a Rex+ Tools",
    descripcion="Selecciona una herramienta para comenzar. Usa el menú lateral para navegar entre las apps disponibles.",
    icono="🛠️",
)

# ─────────────────────────────────────────────
#  DASHBOARD DE HERRAMIENTAS
# ─────────────────────────────────────────────

st.markdown("### Herramientas disponibles")
st.markdown("")

col1, col2 = st.columns(2)

with col1:
    st.markdown(
        '<div class="rex-tool-card">'
        '<div class="rex-tool-icon">📋</div>'
        '<div class="rex-tool-title">Validador de Empleados</div>'
        '<div class="rex-tool-desc">'
        'Valida y corrige automáticamente archivos Excel de empleados: RUT, teléfonos, '
        'fechas, comunas, regiones y ciudades según el maestro oficial.'
        '</div>'
        '<span class="rex-tool-tag activa">ACTIVA</span>'
        '<span class="rex-tool-tag">Excel</span>'
        '<span class="rex-tool-tag">Chile</span>'
        '<br/><br/>'
        '<span class="rex-tool-cta">Abrir herramienta →</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    if st.button("Abrir Validador", key="btn_validador", use_container_width=True):
        st.switch_page("pages/1_Validador_Empleados.py")

with col2:
    st.markdown(
        '<div class="rex-tool-card">'
        '<div class="rex-tool-icon lime">📄</div>'
        '<div class="rex-tool-title">Progestion - Carga de Documentos</div>'
        '<div class="rex-tool-desc">'
        'Procesa ZIP con PDFs organizados por RUT y genera el archivo '
        'de configuración CSV listo para cargar al sistema de gestión documental.'
        '</div>'
        '<span class="rex-tool-tag activa">ACTIVA</span>'
        '<span class="rex-tool-tag">PDF</span>'
        '<span class="rex-tool-tag">ZIP</span>'
        '<br/><br/>'
        '<span class="rex-tool-cta">Abrir herramienta →</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    if st.button("Abrir Progestion", key="btn_progestion", use_container_width=True):
        st.switch_page("pages/2_Progestion.py")

# Placeholder para futuras herramientas
st.markdown("")
col3, col4 = st.columns(2)

with col3:
    st.markdown(
        '<div class="rex-tool-card" style="opacity:0.6;">'
        '<div class="rex-tool-icon dark">⚙️</div>'
        '<div class="rex-tool-title">Próximamente</div>'
        '<div class="rex-tool-desc">'
        'Espacio reservado para la próxima herramienta Rex+. '
        'Si tienes una idea o necesitas automatizar algún proceso, ¡pídelo!'
        '</div>'
        '<span class="rex-tool-tag proximo">PRÓXIMAMENTE</span>'
        '</div>',
        unsafe_allow_html=True,
    )

# ─────────────────────────────────────────────
#  ESTADÍSTICAS RÁPIDAS
# ─────────────────────────────────────────────
st.markdown("---")
st.markdown("### Resumen")

m1, m2, m3, m4 = st.columns(4)
m1.metric("🛠️ Herramientas", "2")
m2.metric("✅ Activas", "2")
m3.metric("🔜 En desarrollo", "0")
m4.metric("📅 Versión", "1.0")

aplicar_footer()
