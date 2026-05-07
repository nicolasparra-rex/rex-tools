import sys
from pathlib import Path

# Agregar la carpeta acta_app al path para importar sus módulos
sys.path.insert(0, str(Path(__file__).parent.parent / "acta_app"))

# Ejecutar la app directamente
exec(open(Path(__file__).parent.parent / "acta_app" / "app.py").read())
