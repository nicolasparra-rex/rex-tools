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

st.markdown("")
col3, col4 = st.columns(2)

with col3:
    st.markdown(
        '<div class="rex-tool-card">'
        '<div class="rex-tool-icon dark">📂</div>'
        '<div class="rex-tool-title">Gestión Documentos</div>'
        '<div class="rex-tool-desc">'
        'Procesa ZIP con PDFs organizados por carpeta APELLIDO_RUT y genera el CSV '
        'de configuración usando un mapeo dinámico de tipos de documento (Excel cargable).'
        '</div>'
        '<span class="rex-tool-tag activa">ACTIVA</span>'
        '<span class="rex-tool-tag">PDF</span>'
        '<span class="rex-tool-tag">ZIP</span>'
        '<span class="rex-tool-tag">Mapeo</span>'
        '<br/><br/>'
        '<span class="rex-tool-cta">Abrir herramienta →</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    if st.button("Abrir Gestión Documentos", key="btn_gestion_docs", use_container_width=True):
        st.switch_page("pages/3_Gestion_Documentos.py")

with col4:
    st.markdown(
        '<div class="rex-tool-card">'
        '<div class="rex-tool-icon lime">📝</div>'
        '<div class="rex-tool-title">Acta de Implementación</div>'
        '<div class="rex-tool-desc">'
        'Genera actas de implementación automáticamente a partir de los datos '
        'del cliente y configuración del proyecto.'
        '</div>'
        '<span class="rex-tool-tag activa">ACTIVA</span>'
        '<span class="rex-tool-tag">Acta</span>'
        '<span class="rex-tool-tag">Word</span>'
        '<br/><br/>'
        '<span class="rex-tool-cta">Abrir herramienta →</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    if st.button("Abrir Acta Implementación", key="btn_acta", use_container_width=True):
        st.switch_page("pages/4_Acta_Implementacion.py")

st.markdown("")
col5, col6 = st.columns(2)

with col5:
    st.markdown(
        '<div class="rex-tool-card">'
        '<div class="rex-tool-icon">📊</div>'
        '<div class="rex-tool-title">Libro de Remuneraciones</div>'
        '<div class="rex-tool-desc">'
        'Transforma el CSV del Libro de Remuneraciones Electrónico al formato '
        'de importación Rex+. Soporta múltiples archivos en formato RUTempresa_YYYYMM.csv.'
        '</div>'
        '<span class="rex-tool-tag activa">ACTIVA</span>'
        '<span class="rex-tool-tag">CSV</span>'
        '<span class="rex-tool-tag">LRE</span>'
        '<br/><br/>'
        '<span class="rex-tool-cta">Abrir herramienta →</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    if st.button("Abrir Libro Remuneraciones", key="btn_libro_rem", use_container_width=True):
        st.switch_page("pages/5_Libro_Remuneraciones.py")

with col6:
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
m1.metric("🛠️ Herramientas", "5")
m2.metric("✅ Activas", "5")
m3.metric("🔜 En desarrollo", "0")
m4.metric("📅 Versión", "1.3")

aplicar_footer()
