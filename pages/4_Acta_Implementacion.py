import sys
from pathlib import Path
import streamlit as st
 
# Aplicar branding Rex+ estándar
sys.path.insert(0, str(Path(__file__).parent.parent))
from lib.branding import aplicar_branding, aplicar_footer
 
# Neutralizar set_page_config para evitar conflicto con el hub
st.set_page_config = lambda **kwargs: None
 
# Neutralizar el CSS personalizado de acta_app (se reemplaza por el branding Rex+)
_original_markdown = st.markdown
def _markdown_sin_estilos(body, **kwargs):
    if isinstance(body, str) and "<style>" in body:
        return  # ignorar bloques de estilo propios de acta_app
    return _original_markdown(body, **kwargs)
st.markdown = _markdown_sin_estilos
 
# Aplicar branding Rex+
aplicar_branding(titulo_pagina="Acta de Implementación", badge="PRODUCCION")
 
# Restaurar st.markdown para el resto de la app
st.markdown = _original_markdown
 
# Agregar la carpeta acta_app al path para importar sus módulos
sys.path.insert(0, str(Path(__file__).parent.parent / "acta_app"))
 
# Ejecutar la app de acta
exec(open(Path(__file__).parent.parent / "acta_app" / "app.py").read())
