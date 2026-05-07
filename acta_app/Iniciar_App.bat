@echo off
cd /d "%~dp0"
echo Iniciando Acta de Implementacion REX+...
echo.
echo Cuando veas el mensaje "You can now view...", abre tu navegador en:
echo http://localhost:8501
echo.
python -m streamlit run app.py --server.headless true
