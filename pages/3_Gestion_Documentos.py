"""
Gestión Documentos - Procesamiento de ZIPs con mapeo dinámico de tipos
========================================================================
Procesa un ZIP con PDFs organizados por carpetas (APELLIDO_RUT/)
y genera configuracion.csv usando un Excel de mapeo (palabra clave → código)
que puede ser default del repo o subido por el usuario.

Estructura esperada del ZIP:
    APELLIDO_RUT/
    ├── APELLIDO_RUT_1_CONTRATO DE TRABAJO.pdf
    ├── APELLIDO_RUT_1_Licencia Médica 2024.pdf
    └── APELLIDO_RUT_1_PROTOCOLOS.pdf
"""

import streamlit as st
import pandas as pd
import io
import csv
import re
import zipfile
import unicodedata
from datetime import datetime
from pathlib import Path, PurePosixPath

from lib.branding import aplicar_branding, aplicar_footer, hero

st.set_page_config(
    page_title="Gestión Documentos | Rex+ Tools",
    page_icon="📂",
    layout="wide",
)

aplicar_branding(titulo_pagina="Gestión Documentos")

hero(
    titulo="Gestión Documentos",
    descripcion="Sube un ZIP con PDFs organizados por RUT y un Excel de mapeo (palabra clave → código) para generar el configuracion.csv.",
    icono="📂",
)


# ─────────────────────────────────────────────
#  CONFIGURACIÓN
# ─────────────────────────────────────────────
OPCION_EMITIR  = "N"
CONTRATO       = "1"
CODIGO_NO_DEFINIDO = "ND"

COLUMNAS_CSV = [
    "Rut del empleado",
    "Contrato",
    "Tipo de documento",
    "Nombre del nuevo documento",
    "Nombre archivo con extensión",
    "Opción emitir",
]

# Ruta al mapeo default del repo
RUTA_MAPEO_DEFAULT = Path(__file__).parent.parent / "assets" / "mapeo_default.xlsx"


# ─────────────────────────────────────────────
#  UTILIDADES
# ─────────────────────────────────────────────
def _normalizar(texto):
    """Minúsculas, sin tildes, espacios limpios."""
    if texto is None:
        return ""
    s = str(texto).lower().strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s


def formatear_rut(rut):
    """Normaliza el RUT al formato 12345678-9, quitando ceros a la izquierda."""
    if not rut:
        return rut
    rut = str(rut).strip().upper().replace(".", "").replace(" ", "")
    rut_sin_guion = rut.replace("-", "")
    if len(rut_sin_guion) < 2:
        return rut
    cuerpo = rut_sin_guion[:-1].lstrip("0")
    dv = rut_sin_guion[-1]
    return f"{cuerpo}-{dv}"


def _corregir_filename(nombre):
    """Corrige el encoding de nombres en ZIP (Mac UTF-8 NFD o Windows CP1252)."""
    try:
        raw = nombre.encode("cp437")
        try:
            return raw.decode("utf-8")
        except UnicodeDecodeError:
            return raw.decode("cp1252", errors="replace")
    except UnicodeEncodeError:
        return nombre


def cargar_mapeo(archivo_excel):
    """Carga el Excel de mapeo y retorna una lista ordenada por longitud descendente.

    El Excel debe tener columnas: 'Palabra clave', 'Código', 'Descripción'
    """
    df = pd.read_excel(archivo_excel, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    # Validar columnas
    esperadas = {"Palabra clave", "Código", "Descripción"}
    if not esperadas.issubset(set(df.columns)):
        raise ValueError(
            f"El Excel debe tener las columnas: {esperadas}. "
            f"Encontradas: {set(df.columns)}"
        )

    # Construir lista (palabra_normalizada, código, descripción, palabra_original)
    mapeo = []
    for _, fila in df.iterrows():
        palabra = fila["Palabra clave"]
        codigo  = fila["Código"]
        desc    = fila["Descripción"]
        if pd.isna(palabra) or pd.isna(codigo) or not str(palabra).strip():
            continue
        palabra_norm = _normalizar(palabra)
        mapeo.append((palabra_norm, str(codigo).strip(), str(desc).strip(), str(palabra).strip()))

    # Ordenar por longitud descendente → palabras más específicas primero
    mapeo.sort(key=lambda x: -len(x[0]))
    return mapeo


def buscar_tipo(nombre_archivo, mapeo):
    """Busca el tipo de documento en el mapeo según el nombre del archivo.

    Retorna (codigo, descripcion, palabra_matcheada) o ('ND', 'No Definido', None).
    """
    nombre_norm = _normalizar(nombre_archivo)
    for keyword, codigo, desc, original in mapeo:
        if keyword and keyword in nombre_norm:
            return codigo, desc, original
    return CODIGO_NO_DEFINIDO, "No Definido", None


def extraer_rut_de_carpeta(nombre_carpeta):
    """Extrae el RUT desde el nombre de la carpeta (ej: 'ABREU_26860900-8' → '26860900-8')."""
    # Buscar un patrón que parezca RUT: dígitos + opcional guión + dígito/K al final
    match = re.search(r"(\d{6,9}[-]?[\dkK])$", nombre_carpeta.strip())
    if match:
        return formatear_rut(match.group(1))
    # Si no se encuentra, intentar última parte después del último _
    if "_" in nombre_carpeta:
        ultima_parte = nombre_carpeta.split("_")[-1]
        return formatear_rut(ultima_parte)
    return None


def extraer_tipo_desde_nombre_pdf(nombre_pdf):
    """Extrae la parte del nombre del PDF que sirve para identificar el tipo.

    Formatos esperados:
    - 'APELLIDO_RUT_N_TIPO.pdf' → retorna 'TIPO'
    - 'TIPO.pdf' → retorna 'TIPO'
    """
    # Quitar extensión
    sin_ext = PurePosixPath(nombre_pdf).stem

    # Si tiene patrón APELLIDO_RUT_N_TIPO, quedarnos con todo después del 3er _
    partes = sin_ext.split("_")
    if len(partes) >= 4:
        # Los 3 primeros son APELLIDO, RUT, N → el resto es el tipo
        return "_".join(partes[3:])

    # Si no tiene ese patrón, retornar el nombre completo
    return sin_ext


def procesar_zip(archivo_zip, mapeo):
    """Procesa el ZIP y retorna (filas_csv, errores, stats)."""
    filas = []
    errores = []
    sin_match = []  # nombres que cayeron en ND
    stats = {
        "total_pdfs":     0,
        "ruts_unicos":    set(),
        "pdfs_ignorados": 0,
        "sin_match":      0,
        "codigos_usados": {},  # conteo por código
    }

    with zipfile.ZipFile(archivo_zip, "r") as zf:
        nombres_zip = zf.namelist()

        for nombre in nombres_zip:
            nombre_limpio = _corregir_filename(nombre)

            if nombre_limpio.endswith("/"):
                continue

            base = PurePosixPath(nombre_limpio).name
            if base.startswith(".") or base.startswith("__MACOSX") or base == "Thumbs.db":
                stats["pdfs_ignorados"] += 1
                continue

            if not base.lower().endswith(".pdf"):
                stats["pdfs_ignorados"] += 1
                continue

            partes = PurePosixPath(nombre_limpio).parts
            if "__MACOSX" in partes:
                stats["pdfs_ignorados"] += 1
                continue

            if len(partes) < 2:
                errores.append(f"PDF sin carpeta padre: {base}")
                continue

            # Extraer RUT desde la carpeta padre
            carpeta_rut = partes[-2]
            rut = extraer_rut_de_carpeta(carpeta_rut)
            if not rut:
                errores.append(f"No se pudo extraer RUT de carpeta '{carpeta_rut}' para archivo {base}")
                continue

            # Extraer la parte del nombre que define el tipo
            sin_ext = PurePosixPath(base).stem
            tipo_texto = extraer_tipo_desde_nombre_pdf(base)

            # Buscar en el mapeo
            codigo, descripcion, palabra = buscar_tipo(tipo_texto, mapeo)

            # También probar con el nombre completo (por si acaso)
            if codigo == CODIGO_NO_DEFINIDO:
                codigo2, desc2, palabra2 = buscar_tipo(sin_ext, mapeo)
                if codigo2 != CODIGO_NO_DEFINIDO:
                    codigo, descripcion, palabra = codigo2, desc2, palabra2

            if codigo == CODIGO_NO_DEFINIDO:
                stats["sin_match"] += 1
                sin_match.append({"rut": rut, "archivo": base, "tipo_extraido": tipo_texto})

            filas.append({
                "Rut del empleado":            rut,
                "Contrato":                    CONTRATO,
                "Tipo de documento":           codigo,
                "Nombre del nuevo documento":  sin_ext,
                "Nombre archivo con extensión": base,
                "Opción emitir":               OPCION_EMITIR,
            })

            stats["total_pdfs"] += 1
            stats["ruts_unicos"].add(rut)
            stats["codigos_usados"][codigo] = stats["codigos_usados"].get(codigo, 0) + 1

    return filas, errores, stats, sin_match


def generar_csv(filas):
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

# ─── Instrucciones ───
st.markdown("### 📋 Cómo funciona")

col_i1, col_i2, col_i3 = st.columns(3)

with col_i1:
    st.markdown(
        '<div style="background:white;border:1px solid #E8EEF3;border-radius:12px;padding:1.25rem;height:100%;">'
        '<div style="color:#1EBBEF;font-weight:700;font-size:0.75rem;letter-spacing:0.5px;margin-bottom:0.5rem;">PASO 1</div>'
        '<div style="color:#1A3A5F;font-size:1rem;font-weight:600;margin-bottom:0.25rem;">Mapeo de tipos</div>'
        '<div style="color:#8B9DAE;font-size:0.8rem;">Sube un Excel con columnas <strong>Palabra clave, Código, Descripción</strong> (o usa el default)</div>'
        '</div>',
        unsafe_allow_html=True,
    )

with col_i2:
    st.markdown(
        '<div style="background:white;border:1px solid #E8EEF3;border-radius:12px;padding:1.25rem;height:100%;">'
        '<div style="color:#1EBBEF;font-weight:700;font-size:0.75rem;letter-spacing:0.5px;margin-bottom:0.5rem;">PASO 2</div>'
        '<div style="color:#1A3A5F;font-size:1rem;font-weight:600;margin-bottom:0.25rem;">ZIP con PDFs</div>'
        '<div style="color:#8B9DAE;font-size:0.8rem;">Sube un ZIP con carpetas <strong>APELLIDO_RUT/</strong> que contengan los PDFs</div>'
        '</div>',
        unsafe_allow_html=True,
    )

with col_i3:
    st.markdown(
        '<div style="background:white;border:1px solid #E8EEF3;border-radius:12px;padding:1.25rem;height:100%;">'
        '<div style="color:#1EBBEF;font-weight:700;font-size:0.75rem;letter-spacing:0.5px;margin-bottom:0.5rem;">PASO 3</div>'
        '<div style="color:#1A3A5F;font-size:1rem;font-weight:600;margin-bottom:0.25rem;">Descargar CSV</div>'
        '<div style="color:#8B9DAE;font-size:0.8rem;">La app genera el <strong>configuracion.csv</strong> listo para importar</div>'
        '</div>',
        unsafe_allow_html=True,
    )

st.markdown("")

# ─── Uploaders ───
st.markdown("### 📤 Subir archivos")

col_u1, col_u2 = st.columns(2)

with col_u1:
    st.markdown("**📊 Mapeo de tipos de documento**")
    archivo_mapeo = st.file_uploader(
        "Excel con columnas: Palabra clave, Código, Descripción",
        type=["xlsx", "xls"],
        key="mapeo",
        help="Opcional. Si no subes uno, se usa el mapeo default del repo.",
    )

    # Guardar en session_state para cachear entre sesiones de la misma pestaña
    if archivo_mapeo is not None:
        st.session_state["ultimo_mapeo"] = archivo_mapeo.getvalue()
        st.session_state["ultimo_mapeo_nombre"] = archivo_mapeo.name
        st.success(f"✅ Mapeo cargado: {archivo_mapeo.name}")
    elif "ultimo_mapeo" in st.session_state:
        st.info(f"📎 Usando mapeo anterior: {st.session_state['ultimo_mapeo_nombre']}")
    else:
        if RUTA_MAPEO_DEFAULT.exists():
            st.info("📎 Usando mapeo default del repo")
        else:
            st.warning("⚠️ No hay mapeo default, sube uno para continuar")

with col_u2:
    st.markdown("**📦 ZIP con PDFs**")
    archivo_zip = st.file_uploader(
        "ZIP con carpetas APELLIDO_RUT/",
        type=["zip"],
        key="zip",
        help="Cada carpeta debe llamarse 'APELLIDO_RUT' y contener los PDFs del empleado.",
    )

# ─── Procesamiento ───
if archivo_zip is not None:
    # Resolver qué mapeo usar
    try:
        if archivo_mapeo is not None:
            # Recién subido
            mapeo = cargar_mapeo(archivo_mapeo)
            origen_mapeo = f"Excel subido: {archivo_mapeo.name}"
        elif "ultimo_mapeo" in st.session_state:
            # Cache de sesión anterior
            mapeo = cargar_mapeo(io.BytesIO(st.session_state["ultimo_mapeo"]))
            origen_mapeo = f"Cache: {st.session_state['ultimo_mapeo_nombre']}"
        elif RUTA_MAPEO_DEFAULT.exists():
            # Default del repo
            mapeo = cargar_mapeo(str(RUTA_MAPEO_DEFAULT))
            origen_mapeo = "Mapeo default del repo"
        else:
            st.error("❌ No hay mapeo disponible. Sube un Excel de mapeo para continuar.")
            st.stop()
    except Exception as e:
        st.error(f"❌ Error leyendo el Excel de mapeo: {e}")
        st.stop()

    with st.spinner("Procesando ZIP..."):
        try:
            filas, errores, stats, sin_match = procesar_zip(archivo_zip, mapeo)

            if not filas:
                st.warning("⚠️ No se encontraron PDFs válidos en el archivo ZIP.")
            else:
                st.success(f"✅ Procesamiento completado")
                st.caption(f"Mapeo usado: {origen_mapeo} · {len(mapeo)} palabras clave")

                # ─── Métricas ───
                st.markdown("### Resumen del proceso")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("📄 PDFs procesados", stats["total_pdfs"])
                c2.metric("👥 RUTs únicos", len(stats["ruts_unicos"]))
                c3.metric("🚫 Archivos ignorados", stats["pdfs_ignorados"])
                c4.metric("❓ Sin match (ND)", stats["sin_match"])

                # ─── Top códigos usados ───
                if stats["codigos_usados"]:
                    with st.expander("📊 Ver distribución por tipo de documento", expanded=False):
                        df_codigos = pd.DataFrame([
                            {"Código": k, "Cantidad": v}
                            for k, v in sorted(stats["codigos_usados"].items(), key=lambda x: -x[1])
                        ])
                        st.dataframe(df_codigos, use_container_width=True, hide_index=True)

                # ─── Archivos sin match ───
                if sin_match:
                    with st.expander(f"❓ Ver {len(sin_match)} archivo(s) sin categoría (código ND)", expanded=False):
                        st.caption("Estos archivos no matchean ninguna palabra clave del mapeo. Se asignaron al código 'ND' (No Definido).")
                        df_sin_match = pd.DataFrame(sin_match)
                        st.dataframe(df_sin_match, use_container_width=True, hide_index=True)

                # ─── Vista previa ───
                st.markdown("### Vista previa del CSV")
                df_preview = pd.DataFrame(filas)
                st.dataframe(df_preview.head(20), use_container_width=True)

                if len(filas) > 20:
                    st.caption(f"Mostrando 20 de {len(filas)} filas totales.")

                # ─── Descarga ───
                st.markdown("### Descargar archivo")

                contenido_csv = generar_csv(filas)

                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    st.download_button(
                        label="📥 Descargar configuracion.csv",
                        data=contenido_csv.encode("utf-8-sig"),
                        file_name="configuracion.csv",
                        mime="text/csv",
                        use_container_width=True,
                        type="primary",
                    )

                # ─── Errores ───
                if errores:
                    with st.expander(f"⚠️ Ver {len(errores)} error(es) detectados"):
                        for err in errores:
                            st.markdown(f"- {err}")

                    reporte_txt = (
                        f"REPORTE DE PROCESO - Gestión Documentos\n"
                        f"{'=' * 60}\n\n"
                        f"Fecha: {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}\n"
                        f"Mapeo usado: {origen_mapeo}\n"
                        f"PDFs procesados: {stats['total_pdfs']}\n"
                        f"RUTs únicos: {len(stats['ruts_unicos'])}\n"
                        f"Archivos ignorados: {stats['pdfs_ignorados']}\n"
                        f"Sin match (ND): {stats['sin_match']}\n"
                        f"Errores: {len(errores)}\n\n"
                        f"DETALLE DE ERRORES\n"
                        f"{'-' * 40}\n"
                    )
                    for err in errores:
                        reporte_txt += f"- {err}\n"

                    with col_d2:
                        st.download_button(
                            label="📄 Descargar reporte de errores",
                            data=reporte_txt.encode("utf-8"),
                            file_name="reporte_errores.txt",
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
