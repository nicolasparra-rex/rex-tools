"""
Rex+ Tools - Página principal
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

st.markdown("### Herramientas disponibles")
st.markdown("")

# ── Fila 1 ────────────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.markdown(
        '<div class="rex-tool-card">'
        '<div class="rex-tool-icon">📋</div>'
        '<div class="rex-tool-title">1 · Validador de Empleados</div>'
        '<div class="rex-tool-desc">'
        'Valida y corrige automáticamente archivos Excel de empleados: RUT, teléfonos, '
        'fechas, comunas, regiones y ciudades según el maestro oficial.'
        '</div>'
        '<span class="rex-tool-tag activa">ACTIVA</span>'
        '<span class="rex-tool-tag">Excel</span>'
        '<span class="rex-tool-tag">Chile</span>'
        '<br/><br/><span class="rex-tool-cta">Abrir herramienta →</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    if st.button("Abrir Validador", key="btn_validador", use_container_width=True):
        st.switch_page("pages/1_Validador_Empleados.py")

with col2:
    st.markdown(
        '<div class="rex-tool-card">'
        '<div class="rex-tool-icon lime">📄</div>'
        '<div class="rex-tool-title">2 · Progestion - Carga de Documentos</div>'
        '<div class="rex-tool-desc">'
        'Procesa ZIP con PDFs organizados por RUT y genera el archivo '
        'de configuración CSV listo para cargar al sistema de gestión documental.'
        '</div>'
        '<span class="rex-tool-tag activa">ACTIVA</span>'
        '<span class="rex-tool-tag">PDF</span>'
        '<span class="rex-tool-tag">ZIP</span>'
        '<br/><br/><span class="rex-tool-cta">Abrir herramienta →</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    if st.button("Abrir Progestion", key="btn_progestion", use_container_width=True):
        st.switch_page("pages/2_Progestion.py")

st.markdown("")

# ── Fila 2 ────────────────────────────────────────────────────────────────────
col3, col4 = st.columns(2)

with col3:
    st.markdown(
        '<div class="rex-tool-card">'
        '<div class="rex-tool-icon dark">📂</div>'
        '<div class="rex-tool-title">3 · Gestión Documentos</div>'
        '<div class="rex-tool-desc">'
        'Procesa ZIP con PDFs organizados por carpeta APELLIDO_RUT y genera el CSV '
        'de configuración usando un mapeo dinámico de tipos de documento (Excel cargable).'
        '</div>'
        '<span class="rex-tool-tag activa">ACTIVA</span>'
        '<span class="rex-tool-tag">PDF</span>'
        '<span class="rex-tool-tag">ZIP</span>'
        '<span class="rex-tool-tag">Mapeo</span>'
        '<br/><br/><span class="rex-tool-cta">Abrir herramienta →</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    if st.button("Abrir Gestión Documentos", key="btn_gestion_docs", use_container_width=True):
        st.switch_page("pages/3_Gestion_Documentos.py")

with col4:
    st.markdown(
        '<div class="rex-tool-card">'
        '<div class="rex-tool-icon lime">📝</div>'
        '<div class="rex-tool-title">4 · Acta de Implementación</div>'
        '<div class="rex-tool-desc">'
        'Genera actas de implementación automáticamente a partir de los datos '
        'del cliente y configuración del proyecto. Búsqueda por OT desde Zoho.'
        '</div>'
        '<span class="rex-tool-tag activa">ACTIVA</span>'
        '<span class="rex-tool-tag">Acta</span>'
        '<span class="rex-tool-tag">Word</span>'
        '<span class="rex-tool-tag">Zoho</span>'
        '<br/><br/><span class="rex-tool-cta">Abrir herramienta →</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    if st.button("Abrir Acta Implementación", key="btn_acta", use_container_width=True):
        st.switch_page("pages/4_Acta_Implementacion.py")

st.markdown("")

# ── Fila 3 ────────────────────────────────────────────────────────────────────
col5, col6 = st.columns(2)

with col5:
    st.markdown(
        '<div class="rex-tool-card">'
        '<div class="rex-tool-icon">📊</div>'
        '<div class="rex-tool-title">5 · Libro de Remuneraciones</div>'
        '<div class="rex-tool-desc">'
        'Transforma el CSV del Libro de Remuneraciones Electrónico al formato '
        'de importación Rex+. Soporta múltiples archivos en formato RUTempresa_YYYYMM.csv.'
        '</div>'
        '<span class="rex-tool-tag activa">ACTIVA</span>'
        '<span class="rex-tool-tag">CSV</span>'
        '<span class="rex-tool-tag">LRE</span>'
        '<br/><br/><span class="rex-tool-cta">Abrir herramienta →</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    if st.button("Abrir Libro Remuneraciones", key="btn_libro_rem", use_container_width=True):
        st.switch_page("pages/5_Libro_Remuneraciones.py")

with col6:
    st.markdown(
        '<div class="rex-tool-card">'
        '<div class="rex-tool-icon lime">📋</div>'
        '<div class="rex-tool-title">6 · Minutas de Implementación</div>'
        '<div class="rex-tool-desc">'
        'Completa los datos del cliente y descarga la minuta en Excel lista para entregar. '
        'Autocompletado desde Zoho al ingresar la OT. Incluye Remuneraciones y Asistencia.'
        '</div>'
        '<span class="rex-tool-tag activa">ACTIVA</span>'
        '<span class="rex-tool-tag">Excel</span>'
        '<span class="rex-tool-tag">Zoho</span>'
        '<br/><br/><span class="rex-tool-cta">Abrir herramienta →</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    if st.button("Abrir Minutas", key="btn_minutas", use_container_width=True):
        st.switch_page("pages/7_Minutas.py")

st.markdown("")

# ── Fila 4 ────────────────────────────────────────────────────────────────────
col7, col8 = st.columns(2)

with col7:
    st.markdown(
        '<div class="rex-tool-card">'
        '<div class="rex-tool-icon dark">🗂️</div>'
        '<div class="rex-tool-title">7 · Zoho Proyectos</div>'
        '<div class="rex-tool-desc">'
        'Visualiza todos los proyectos activos del portal Rex+ en tiempo real. '
        'Filtros por estado, plan, consultor y grupo. Detalle de tareas por proyecto.'
        '</div>'
        '<span class="rex-tool-tag activa">ACTIVA</span>'
        '<span class="rex-tool-tag">Zoho</span>'
        '<span class="rex-tool-tag">Tiempo real</span>'
        '<br/><br/><span class="rex-tool-cta">Abrir herramienta →</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    if st.button("Abrir Zoho Proyectos", key="btn_zoho", use_container_width=True):
        st.switch_page("pages/8_Zoho_Proyectos.py")

with col8:
    st.markdown(
        '<div class="rex-tool-card">'
        '<div class="rex-tool-icon orange">📊</div>'
        '<div class="rex-tool-title">8 · Dashboard</div>'
        '<div class="rex-tool-desc">'
        'KPIs de cartera en tiempo real: activos, detenidos, salidas planificadas '
        'por mes y proyectos por consultor, grupo y plan.'
        '</div>'
        '<span class="rex-tool-tag activa">ACTIVA</span>'
        '<span class="rex-tool-tag">KPIs</span>'
        '<span class="rex-tool-tag">Zoho</span>'
        '<br/><br/><span class="rex-tool-cta">Abrir herramienta →</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    if st.button("Abrir Dashboard", key="btn_dashboard", use_container_width=True):
        st.switch_page("pages/9_Dashboard.py")

st.markdown("")

# ── Fila 5 ────────────────────────────────────────────────────────────────────
col9, col10 = st.columns(2)

with col9:
    st.markdown(
        '<div class="rex-tool-card">'
        '<div class="rex-tool-icon">📑</div>'
        '<div class="rex-tool-title">6 · LRE Detalle</div>'
        '<div class="rex-tool-desc">'
        'Procesamiento y detalle del Libro de Remuneraciones Electrónico. '
        'Visualiza y analiza la información por período y empresa.'
        '</div>'
        '<span class="rex-tool-tag activa">ACTIVA</span>'
        '<span class="rex-tool-tag">LRE</span>'
        '<span class="rex-tool-tag">Remuneraciones</span>'
        '<br/><br/><span class="rex-tool-cta">Abrir herramienta →</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    if st.button("Abrir LRE Detalle", key="btn_lre_detalle", use_container_width=True):
        st.switch_page("pages/6_Lre_Detalle.py")

with col10:
    st.markdown(
        '<div class="rex-tool-card">'
        '<div class="rex-tool-icon orange">🔄</div>'
        '<div class="rex-tool-title">10 · App Migración</div>'
        '<div class="rex-tool-desc">'
        'Herramienta de migración de datos entre sistemas. '
        'Facilita el traspaso y transformación de información al formato Rex+.'
        '</div>'
        '<span class="rex-tool-tag activa">ACTIVA</span>'
        '<span class="rex-tool-tag">Migración</span>'
        '<span class="rex-tool-tag">Datos</span>'
        '<br/><br/><span class="rex-tool-cta">Abrir herramienta →</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    if st.button("Abrir App Migración", key="btn_migracion", use_container_width=True):
        st.switch_page("pages/10_App_Migracion.py")
st.markdown("---")
st.markdown("### Resumen")

m1, m2, m3, m4 = st.columns(4)
m1.metric("🛠️ Herramientas", "10")
m2.metric("✅ Activas", "10")
m3.metric("🔜 En desarrollo", "0")
m4.metric("📅 Versión", "1.5")

aplicar_footer()
