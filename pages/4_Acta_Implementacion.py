import sys
from pathlib import Path
import streamlit as st

# Agregar la carpeta acta_app al path para importar sus módulos
sys.path.insert(0, str(Path(__file__).parent.parent / "acta_app"))

# Neutralizar set_page_config pero mantener layout="wide"
st.set_page_config = lambda **kwargs: None

# Ejecutar la app directamente
exec(open(Path(__file__).parent.parent / "acta_app" / "app.py").read())
