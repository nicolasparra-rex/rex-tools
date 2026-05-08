import sys
from pathlib import Path
import streamlit as st

# Neutralizar set_page_config para evitar conflicto con el hub
st.set_page_config = lambda **kwargs: None

# Neutralizar solo el header HTML de acta_app
_original_markdown = st.markdown
def _filtrar_header(body, **kwargs):
    if isinstance(body, str) and "rex-header" in body:
        return
    return _original_markdown(body, **kwargs)
st.markdown = _filtrar_header

# Agregar la carpeta acta_app al path
sys.path.insert(0, str(Path(__file__).parent.parent / "acta_app"))

# Ejecutar la app
exec(open(Path(__file__).parent.parent / "acta_app" / "app.py").read())
