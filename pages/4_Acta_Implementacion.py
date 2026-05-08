import sys
from pathlib import Path
import streamlit as st

# Neutralizar set_page_config para evitar conflicto con el hub
st.set_page_config = lambda **kwargs: None

# Agregar la carpeta acta_app al path para importar sus módulos
sys.path.insert(0, str(Path(__file__).parent.parent / "acta_app"))

# Ejecutar la app directamente
exec(open(Path(__file__).parent.parent / "acta_app" / "app.py").read())
