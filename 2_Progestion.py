"""
Progestion - Carga de Documentos
================================
Procesa un ZIP con PDFs organizados por carpetas (cada carpeta = RUT)
y genera el archivo configuracion.csv para cargar al sistema.

Estructura del ZIP:
    CARPETA_BASE/
    ├── 10067503K/
    │   ├── documento1.pdf
    │   └── documento2.pdf
    ├── 12345678-9/
    │   └── documento3.pdf
    └── ...
"""

import streamlit as st
import pandas as pd
import io
import csv
import zipfile
from datetime import datetime
from pathlib import Path, PurePosixPath

from lib.branding import aplicar_branding, aplicar_footer, hero

st.set_page_config(
    page_title="Progestion | Rex+ Tools",
    page_icon="📄",
    layout="wide",
)

aplicar_branding(titulo_pagina="Progestion")

hero(
    titulo="Progestion - Carga de Documentos",
    descripcion="Sube un archivo ZIP con subcarpetas por RUT y obtén el configuracion.csv listo para cargar.",
    icono="📄",
)


# ─────────────────────────────────────────────
#  CONFIGURACIÓN
# ─────────────────────────────────────────────
TIPO_DOCUMENTO = "HQB"
OPCION_EMITIR  = "N"
CONTRATO       = "1"

COLUMNAS_CSV = [
    "Rut del empleado",
    "Contrato",
    "Tipo de documento",
    "Nombre del nuevo documento",
    "Nombre archivo con extensión",
    "Opción emitir",
]


# ─────────────────────────────────────────────
#  UTILIDADES
# ─────────────────────────────────────────────
def formatear_rut(rut: str) -> str:
    """Normaliza el RUT al formato 12345678-9, quitando ceros a la izquierda."""
    if not rut:
        return rut
    rut = rut.strip().upper().replace(".", "").replace(" ", "")
    # Quitar guión si lo trae
    rut_sin_guion = rut.replace("-", "")
    if len(rut_sin_guion) < 2:
        return rut
    cuerpo = rut_sin_guion[:-1].lstrip("0")
    dv = rut_sin_guion[-1]
    return f"{cuerpo}-{dv}"


def _corregir_filename(nombre: str) -> str:
    """Corrige el encoding de nombres en ZIP (Mac UTF-8 NFD o Windows CP1252)."""
    try:
        # Si el nombre tiene caracteres corruptos, intentar recodificar
        raw = nombre.encode("cp437")
        try:
            return raw.decode("utf-8")
        except UnicodeDecodeError:
            return raw.decode("cp1252", errors="replace")
    except UnicodeEncodeError:
        return nombre


def parsear_nombre_pdf(filename: str) -> tuple[str, str]:
    """Retorna (nombre_sin_extension, nombre_con_extension)."""
    path = PurePosixPath(filename)
    return path.stem, path.name


def procesar_zip(archivo_zip) -> tuple[list[dict], list[str], dict]:
    """Procesa el ZIP y retorna (filas_csv, errores, stats)."""
    filas = []
    errores = []
    stats = {
        "total_pdfs":     0,
        "ruts_unicos":    set(),
        "pdfs_ignorados": 0,
    }

    with zipfile.ZipFile(archivo_zip, "r") as zf:
        nombres_zip = zf.namelist()

        for nombre in nombres_zip:
            nombre_limpio = _corregir_filename(nombre)

            # Ignorar directorios
            if nombre_limpio.endswith("/"):
                continue

            # Ignorar archivos ocultos del sistema (Mac, Windows)
            base = PurePosixPath(nombre_limpio).name
            if base.startswith(".") or base.startswith("__MACOSX") or base == "Thumbs.db":
                stats["pdfs_ignorados"] += 1
                continue

            # Solo considerar PDFs
            if not base.lower().endswith(".pdf"):
                stats["pdfs_ignorados"] += 1
                continue

            # Obtener la carpeta padre del PDF (debe ser el RUT)
            partes = PurePosixPath(nombre_limpio).parts
            if len(partes) < 2:
                errores.append(f"PDF sin carpeta padre: {base}")
                continue

            # La carpeta inmediata que contiene al PDF es el RUT
            carpeta_rut = partes[-2]

            # Ignorar carpeta __MACOSX
            if "__MACOSX" in partes:
                stats["pdfs_ignorados"] += 1
                continue

            rut_formateado = formatear_rut(carpeta_rut)
            if not rut_formateado or len(rut_formateado) < 3:
                errores.append(f"RUT inválido '{carpeta_rut}' en archivo {base}")
                continue

            nombre_sin_ext, nombre_con_ext = parsear_nombre_pdf(base)

            filas.append({
                "Rut del empleado":            rut_formateado,
                "Contrato":                    CONTRATO,
                "Tipo de documento":           TIPO_DOCUMENTO,
                "Nombre del nuevo documento":  nombre_sin_ext,
                "Nombre archivo con extensión": nombre_con_ext,
                "Opción emitir":               OPCION_EMITIR,
            })

            stats["total_pdfs"] += 1
            stats["ruts_unicos"].add(rut_formateado)

    return filas, errores, stats


def generar_csv(filas: list[dict]) -> str:
    """Genera el contenido del CSV a partir de las filas."""
    buffer = io.StringIO()
    writer = csv.DictWriter(buffer, fieldnames=COLUMNAS_CSV)
    writer.writeheader()
    for fila in filas:
        writer.writerow(fila)
    return buffer.getvalue()


# ─────────────────────────────────────────────
#  UI
# ─────────────────────────────────────────────
st.markdown("### Subir archivo ZIP")

archivo_zip = st.file_uploader(
    "Selecciona el archivo ZIP con los PDFs organizados por carpeta de RUT",
    type=["zip"],
    help="Cada carpeta dentro del ZIP debe llamarse con el RUT del empleado y contener sus PDFs.",
)

st.info(
    "💡 **Estructura esperada del ZIP:**\n\n"
    "- `10067503K/` → documento1.pdf, documento2.pdf\n"
    "- `12345678-9/` → documento3.pdf\n"
    "- etc."
)

if archivo_zip is not None:
    with st.spinner("Procesando archivos..."):
        try:
            filas, errores, stats = procesar_zip(archivo_zip)

            if not filas:
                st.warning("⚠️ No se encontraron PDFs válidos en el archivo ZIP.")
            else:
                st.success(f"✅ Procesamiento completado correctamente")

                # ─── Métricas ───
                st.markdown("### Resumen del proceso")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("📄 PDFs procesados", stats["total_pdfs"])
                c2.metric("👥 RUTs únicos", len(stats["ruts_unicos"]))
                c3.metric("🚫 Archivos ignorados", stats["pdfs_ignorados"])
                c4.metric("⚠️ Errores", len(errores))

                # ─── Vista previa ───
                st.markdown("### Vista previa del CSV")
                df_preview = pd.DataFrame(filas)
                st.dataframe(df_preview.head(20), use_container_width=True)

                if len(filas) > 20:
                    st.caption(f"Mostrando 20 de {len(filas)} filas totales.")

                # ─── Descarga ───
                st.markdown("### Descargar archivo")

                contenido_csv = generar_csv(filas)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

                c_d1, c_d2 = st.columns(2)
                with c_d1:
                    st.download_button(
                        label="⬇️ Descargar configuracion.csv",
                        data=contenido_csv.encode("utf-8"),
                        file_name=f"configuracion_{timestamp}.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )

                # ─── Errores ───
                if errores:
                    with st.expander(f"⚠️ Ver {len(errores)} error(es) detectados"):
                        for err in errores:
                            st.markdown(f"- {err}")

                    reporte_txt = (
                        f"REPORTE DE PROCESO - Progestion\n"
                        f"{'=' * 60}\n\n"
                        f"Fecha: {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}\n"
                        f"PDFs procesados: {stats['total_pdfs']}\n"
                        f"RUTs únicos: {len(stats['ruts_unicos'])}\n"
                        f"Archivos ignorados: {stats['pdfs_ignorados']}\n"
                        f"Errores: {len(errores)}\n\n"
                        f"DETALLE DE ERRORES\n"
                        f"{'-' * 40}\n"
                    )
                    for err in errores:
                        reporte_txt += f"- {err}\n"

                    with c_d2:
                        st.download_button(
                            label="⬇️ Descargar reporte de errores",
                            data=reporte_txt.encode("utf-8"),
                            file_name=f"reporte_errores_{timestamp}.txt",
                            mime="text/plain",
                            use_container_width=True,
                        )

        except zipfile.BadZipFile:
            st.error("❌ El archivo no es un ZIP válido.")
        except Exception as e:
            st.error(f"❌ Error al procesar el archivo: {e}")
            import traceback
            with st.expander("Ver detalle técnico"):
                st.code(traceback.format_exc())

aplicar_footer()
