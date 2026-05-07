# Acta de Implementación · REX+

App local para generar actas de implementación a partir de las Notas de Gemini de Google Meet.

## Instalación

```bash
pip install -r requirements.txt
```

## Uso

```bash
streamlit run app.py
```

## Flujo de uso

1. **Seleccionar cliente** — elige de la lista o deja en "Nuevo cliente"
2. **Completar datos** — empresa, jefe de proyecto cliente y usuario implementador son obligatorios
3. **Subir notas Gemini** — descarga el archivo `.docx` desde Google Drive (carpeta Meet Recordings)
4. **Guardar cliente** — para que aparezca en el autocomplete la próxima vez
5. **Generar acta** — descarga el `.docx` listo para revisar y firmar

## Estructura

```
acta_app/
├── app.py              ← UI Streamlit
├── generator.py        ← generación del .docx
├── extractor.py        ← parseo de notas Gemini
├── template.docx       ← template base del acta
├── clientes.json       ← base de datos local de clientes
└── requirements.txt
```

## clientes.json

Se actualiza automáticamente cada vez que guardas un cliente desde la app.
Puedes editarlo manualmente si necesitas corregir datos.

## Fase 2 — Google Drive (próximamente)

Se agregará `drive_client.py` para conectar con Drive y seleccionar
el archivo de notas directamente desde la app, sin necesidad de descargarlo manualmente.
