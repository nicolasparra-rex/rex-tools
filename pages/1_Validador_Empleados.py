import streamlit as st
import pandas as pd
import re
import unicodedata
from datetime import datetime
import io

from comunas_chile import REGIONES, COMUNAS
from lib.branding import aplicar_branding, aplicar_footer, hero

st.set_page_config(
    page_title="Validador de Empleados | Rex+ Tools",
    page_icon="📋",
    layout="wide",
)


# ─────────────────────────────────────────────
#  BRANDING REX+
# ─────────────────────────────────────────────

aplicar_branding(titulo_pagina="Validador de Empleados")

hero(
    titulo="Validador de Empleados",
    descripcion="Sube el archivo Excel para validar y corregir automáticamente los datos antes de importarlos al sistema.",
    icono="📋",
)


# ─────────────────────────────────────────────
#  CONFIGURACIÓN
# ─────────────────────────────────────────────

campos_obligatorios = [
    "Id empleado", "Situación", "Nombres", "Apellido paterno", "Apellido materno",
    "Sexo", "Fecha de nacimiento", "Estado civil", "Numero de teléfono 1",
    "Numero de teléfono 2", "Comuna", "Ciudad", "Region", "Nombre Calle",
    "Numero Calle", "Departamento", "Id nación", "Email institucional",
    "Email personal", "Nivel de estudio", "Profesión", "Licencia de conducir",
    "Id banco", "Cuenta del banco", "Id forma de pago", "Id AFP",
    "Estado de jubilación", "¿Es expatriado?", "Sistema de pensiones",
    "ID INSTITUCION DE SALUD", "Monto cotizado en la Isapre en UF",
    "Moneda de la cotización", "Tramo de asignación familiar",
    "¿Supervisa otros empleados?", "¿Es un perfil solo aprobador?",
    "Número del contrato", "Tipo del contrato", "Fecha de inicio del contrato",
    "Fecha de término del contrato", "Sueldo base", "Cargo", "Id centro de costo",
    "Id sede donde se desempeña", "¿Realiza trabajo pesado?",
    "Porcentaje de cotización por trabajo pesado", "Id sindicato",
    "¿Jornada parcial?", "Permite ausencias en días inhábiles",
    "Horas de trabajo semanales", "Distribución de jornada",
    "¿Cotiza seguro de cesantía?", "Fecha de incorporación al seguro de cesantía",
    "Id empresa", "Id plantilla grupal", "Causal de término del contrato",
    "Fecha de reconocimiento de vacaciones",
    "Número de meses reconocidos con otro empleador", "Nivel SENCE", "Factor SENCE",
    "Pauta contable", "Agrupación de seguridad", "Área", "¿Descansa domingos?",
    "¿Cotiza previsión y salud?", "Empleado con perfil privado", "Código interno",
    "Talla de ropa", "Talla de zapatos", "Detalle contrato", "Supervisor",
    "Modalidad del contrato", "Turno", "Zona extrema", "Permisos administrativos",
    "Unidad de permisos administrativos", "Categoría INE", "Notas",
    "Centro de distribucion", "Fecha primera renovación", "Fecha segunda renovación",
    "Fecha de inicio de vacaciones"
]

campos_fecha = [
    "Fecha de nacimiento", "Fecha de inicio del contrato",
    "Fecha de término del contrato", "Fecha de incorporación al seguro de cesantía",
    "Fecha de reconocimiento de vacaciones", "Fecha adicional 1", "Fecha adicional 2",
    "Fecha de afiliación a AFP", "Fecha primera renovación",
    "Fecha segunda renovación", "Fecha de inicio de vacaciones"
]

campos_telefono = ["Numero de teléfono 1", "Numero de teléfono 2"]
campos_email    = ["Email institucional", "Email personal"]

estados_civiles_validos = ["S", "C", "V", "D", "U"]

formatos_posibles = [
    "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d",
    "%Y/%m/%d", "%d-%m-%y", "%d/%m/%y", "%m/%d/%Y",
]

TELEFONO_DEFAULT = "56 9 2222 2222"
EMAIL_DEFAULT    = "email@email.com"


# ─────────────────────────────────────────────
#  VALIDACIONES / TRANSFORMACIONES
# ─────────────────────────────────────────────

def _vacio(valor) -> bool:
    """Considera vacío: NaN, None real, string vacío, o strings 'None'/'nan'/'null'/'NaT'."""
    if pd.isna(valor):
        return True
    s = str(valor).strip().lower()
    return s == "" or s in {"none", "nan", "null", "nat", "<na>"}


def reparar_mojibake(texto):
    """Repara caracteres mal codificados (mojibake) como 'AVENDA√ëO' → 'AVENDAÑO'.

    Esto ocurre cuando un archivo guardado en UTF-8 es interpretado como Mac Roman.
    Se detecta por la presencia de caracteres marcadores típicos: √, Ã, ¬.
    """
    if _vacio(texto):
        return texto
    s = str(texto)
    # Solo intentar reparar si hay marcadores típicos de mojibake
    # √ y ¬ son típicos de macroman, Ã es típico de latin-1/windows
    if "√" not in s and "Ã" not in s and "¬" not in s:
        return s
    try:
        return s.encode("macroman").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        try:
            # Fallback: latin-1 → utf-8 (para otros tipos de mojibake de Windows)
            return s.encode("latin-1").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            return s  # No se pudo reparar, dejar como está


def _normalizar_texto(texto) -> str:
    """Quita tildes, pasa a minúsculas y limpia espacios."""
    if _vacio(texto):
        return ""
    s = str(texto).strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s


# Índice inverso: nombre normalizado de comuna → código (construido una sola vez al cargar)
_INDICE_NOMBRE_COMUNA = {}
for _cod, _info in COMUNAS.items():
    _nom_norm = _normalizar_texto(_info["nombre"]).replace(" ", "")
    if _nom_norm and _nom_norm not in _INDICE_NOMBRE_COMUNA:
        _INDICE_NOMBRE_COMUNA[_nom_norm] = _cod


def resolver_comuna(valor):
    """Convierte el valor de Comuna a un código de 5 dígitos si es posible.

    Acepta:
    - Código de 4 o 5 dígitos (ej: '1101', '13114') → aplica zfill
    - Nombre de comuna en texto (ej: 'maipu', 'Las Condes') → busca en el maestro

    Retorna (codigo, cambio_realizado) donde cambio_realizado es un string descriptivo o None.
    """
    if _vacio(valor):
        return valor, None

    original = str(valor).strip()

    # Caso 1: es un número (código)
    if original.isdigit():
        codigo = original.zfill(5)
        if codigo != original and codigo in COMUNAS:
            return codigo, f"Código completado con ceros: '{original}' → '{codigo}'"
        return codigo, None

    # Caso 2: es texto (nombre de comuna)
    nombre_norm = _normalizar_texto(original).replace(" ", "")
    if nombre_norm in _INDICE_NOMBRE_COMUNA:
        codigo = _INDICE_NOMBRE_COMUNA[nombre_norm]
        return codigo, f"Comuna convertida de nombre a código: '{original}' → '{codigo}' ({COMUNAS[codigo]['nombre']})"

    # No se pudo resolver, devolver tal cual (se reportará como error después)
    return original, None


def convertir_fecha(valor):
    if _vacio(valor):
        return valor

    # Si ya viene como datetime de pandas/python, formatear directo
    if isinstance(valor, (pd.Timestamp, datetime)):
        try:
            return valor.strftime("%d-%m-%Y")
        except Exception:
            pass

    valor_str = str(valor).strip()

    # Si viene con timestamp tipo "1972-01-24 00:00:00", quedarse solo con la parte de fecha
    if " " in valor_str:
        valor_str = valor_str.split(" ")[0]

    # Si ya está en el formato deseado, retornar tal cual
    try:
        datetime.strptime(valor_str, "%d-%m-%Y")
        return valor_str
    except ValueError:
        pass

    # Probar los distintos formatos posibles
    for fmt in formatos_posibles:
        try:
            return datetime.strptime(valor_str, fmt).strftime("%d-%m-%Y")
        except ValueError:
            continue

    return valor


def _parsear_fecha(valor):
    """Retorna un datetime o None si no se puede parsear."""
    if _vacio(valor):
        return None
    s = str(valor).strip()
    for fmt in ["%d-%m-%Y"] + formatos_posibles:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def validar_email(valor) -> bool:
    if _vacio(valor):
        return True
    v = str(valor).strip()
    patron = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(patron, v) is not None


def validar_rut_dv(valor) -> bool:
    """Valida el dígito verificador de un RUT chileno (módulo 11)."""
    if _vacio(valor):
        return False
    rut = str(valor).strip().upper().replace(".", "").replace("-", "")
    if len(rut) < 2 or not rut[:-1].isdigit():
        return False
    cuerpo, dv = rut[:-1], rut[-1]

    suma, mult = 0, 2
    for d in reversed(cuerpo):
        suma += int(d) * mult
        mult = mult + 1 if mult < 7 else 2

    resto = 11 - (suma % 11)
    if resto == 11:
        dv_esperado = "0"
    elif resto == 10:
        dv_esperado = "K"
    else:
        dv_esperado = str(resto)

    return dv == dv_esperado


def validar_corregir_id(valor):
    """Mantiene el formato de 9 o 10 dígitos, quitando cero inicial si hay 10."""
    if _vacio(valor):
        return valor
    id_str = str(valor).strip()
    if len(id_str) == 10 and id_str[0] == "0":
        return id_str[1:]
    return id_str


def limpiar_direccion(valor):
    if _vacio(valor):
        return valor
    return re.sub(r"[^a-zA-ZáéíóúÁÉÍÓÚñÑ0-9 ]", "", str(valor).strip())


def convertir_email_minuscula(valor):
    if _vacio(valor):
        return valor
    return str(valor).strip().lower()


def normalizar_telefono(valor):
    """Normaliza a formato '56 9 XXXX XXXX'. Retorna (valor_normalizado, cambio, valido)."""
    if _vacio(valor):
        return TELEFONO_DEFAULT, "completado", True

    original = str(valor).strip()
    # Extraer solo dígitos
    digitos = re.sub(r"\D", "", original)

    # Casos posibles
    if digitos.startswith("56") and len(digitos) == 11:
        # +56 + 9 dígitos → 56 X XXXX XXXX
        formateado = f"56 {digitos[2]} {digitos[3:7]} {digitos[7:]}"
    elif len(digitos) == 9:
        # 9 dígitos sin código país
        formateado = f"56 {digitos[0]} {digitos[1:5]} {digitos[5:]}"
    elif len(digitos) == 8:
        # Fijo antiguo sin código área — asumimos Santiago (2)
        formateado = f"56 2 {digitos[:4]} {digitos[4:]}"
    else:
        # No se pudo normalizar — marcar como inválido
        return original, "sin_cambios", False

    cambio = "ya_ok" if formateado == original else "normalizado"
    return formateado, cambio, True


def _normalizar_region(texto) -> str:
    """Normaliza un nombre de región quitando prefijos comunes."""
    t = _normalizar_texto(texto)
    for prefijo in ["region del ", "region de la ", "region de ", "region ",
                    "r. ", "r.m.", "rm "]:
        if t.startswith(prefijo):
            t = t[len(prefijo):]
            break
    # Casos especiales cortos
    abreviaciones = {
        "rm": "metropolitana de santiago",
        "rm.": "metropolitana de santiago",
        "r.m.": "metropolitana de santiago",
        "metropolitana": "metropolitana de santiago",
    }
    return abreviaciones.get(t.strip(), t.strip())


def corregir_ubicacion(codigo_comuna, region_escrita, ciudad_escrita):
    """Corrige región y ciudad tomando la comuna como fuente de verdad.

    - Region final: siempre código de 2 dígitos con cero (ej: "13", "05")
    - Ciudad final: siempre cod_ciudad del maestro (ej: "LasCondes", "Maipu")

    Retorna una tupla (region_corregida, ciudad_corregida, cambios, error).
    """
    cambios = []

    # Si la comuna viene vacía, no podemos corregir nada
    if _vacio(codigo_comuna):
        return region_escrita, ciudad_escrita, cambios, None

    codigo = str(codigo_comuna).strip().zfill(5)

    # Si la comuna no existe en el maestro, reportar error (no podemos autocorregir)
    if codigo not in COMUNAS:
        return region_escrita, ciudad_escrita, cambios, (
            f"Comuna con código '{codigo}' no existe en el maestro"
        )

    info = COMUNAS[codigo]
    nom_comuna    = info["nombre"]
    cod_region_of = info["cod_region"]      # ej: "13"
    nom_region_of = info["nom_region"]      # ej: "Metropolitana de Santiago"
    cod_ciudad_of = info["cod_ciudad"]      # ej: "LasCondes"

    # ─── Verificar y corregir región (siempre dejar código 2 dígitos) ───
    region_final = cod_region_of  # valor oficial por defecto

    if _vacio(region_escrita):
        cambios.append(f"Región completada: → '{cod_region_of}' (desde Comuna {nom_comuna})")
    else:
        region_str = str(region_escrita).strip()
        region_ok = False

        # Ya es el código correcto
        if region_str.isdigit() and region_str.zfill(2) == cod_region_of:
            region_ok = True
        else:
            # Si empieza con dígitos (formato "XX-nombre"), extraer código
            prefijo = region_str[:2] if region_str[:2].isdigit() else None
            if prefijo and prefijo.zfill(2) == cod_region_of:
                region_ok = True
            else:
                # Comparar por nombre normalizado
                escrita_norm = _normalizar_region(region_str)
                oficial_norm = _normalizar_region(nom_region_of)
                esc_sin_esp  = escrita_norm.replace(" ", "")
                ofi_sin_esp  = oficial_norm.replace(" ", "")
                region_ok = (
                    escrita_norm == oficial_norm
                    or escrita_norm in oficial_norm
                    or oficial_norm in escrita_norm
                    or esc_sin_esp == ofi_sin_esp
                    or esc_sin_esp in ofi_sin_esp
                    or ofi_sin_esp in esc_sin_esp
                )

        # Si la región escrita era válida, igual la normalizamos al código de 2 dígitos
        if region_ok:
            if region_str.zfill(2) != cod_region_of:
                cambios.append(
                    f"Región normalizada a código: '{region_escrita}' → '{cod_region_of}'"
                )
        else:
            cambios.append(
                f"Región corregida: '{region_escrita}' → '{cod_region_of}' "
                f"(porque Comuna es {nom_comuna} [{codigo}])"
            )

    # ─── Verificar y corregir ciudad (siempre dejar cod_ciudad del maestro) ───
    ciudad_final = cod_ciudad_of if cod_ciudad_of else ciudad_escrita

    if cod_ciudad_of:
        if _vacio(ciudad_escrita):
            cambios.append(f"Ciudad completada: → '{cod_ciudad_of}' (desde Comuna {nom_comuna})")
        else:
            ciudad_str = str(ciudad_escrita).strip()
            esc_norm = _normalizar_texto(ciudad_str).replace(" ", "").replace("-", "").replace("_", "")
            of_norm  = _normalizar_texto(cod_ciudad_of).replace(" ", "").replace("-", "").replace("_", "")
            nom_norm = _normalizar_texto(info["nom_ciudad"]).replace(" ", "").replace("-", "").replace("_", "")

            # Aceptar si coincide con código de ciudad oficial o con el nombre (también
            # formatos "05-los_andes_calle_larga" donde el texto después del "-" contiene
            # el nombre de la ciudad)
            ciudad_ok = (
                esc_norm == of_norm
                or esc_norm == nom_norm
                or of_norm in esc_norm
                or nom_norm in esc_norm
                or esc_norm in of_norm
            )

            if ciudad_ok:
                if ciudad_str != cod_ciudad_of:
                    cambios.append(
                        f"Ciudad normalizada: '{ciudad_escrita}' → '{cod_ciudad_of}'"
                    )
            else:
                cambios.append(
                    f"Ciudad corregida: '{ciudad_escrita}' → '{cod_ciudad_of}' "
                    f"(porque Comuna es {nom_comuna} [{codigo}])"
                )

    return region_final, ciudad_final, cambios, None


# ─────────────────────────────────────────────
#  PROCESAMIENTO PRINCIPAL
# ─────────────────────────────────────────────

def procesar_archivo(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name="Empleados", dtype=str)
    total_original = len(df)

    # ───── Reparar caracteres corruptos (mojibake) en headers y celdas ─────
    # Ej: "AVENDA√ëO" → "AVENDAÑO", "tel√©fono" → "teléfono"
    mojibake_reparados = 0
    df.columns = [reparar_mojibake(c) for c in df.columns]
    for col in df.columns:
        if df[col].dtype == object:  # solo columnas de texto
            antes = df[col].copy()
            df[col] = df[col].apply(reparar_mojibake)
            mojibake_reparados += (antes.fillna("") != df[col].fillna("")).sum()

    # Filtrar solo empleados activos
    if "Situación" in df.columns:
        df = df[df["Situación"].str.strip().str.upper() == "A"].reset_index(drop=True)
    filas_eliminadas = total_original - len(df)

    # Contador de correcciones
    correcciones = {
        "fechas_normalizadas":       0,
        "ids_corregidos":            0,
        "comunas_rellenadas":        0,
        "emails_minuscula":          0,
        "emails_vacios_completados": 0,
        "telefonos_normalizados":    0,
        "telefonos_vacios_completados": 0,
        "direcciones_limpiadas":     0,
        "regiones_corregidas":       0,
        "ciudades_corregidas":       0,
        "caracteres_reparados":      int(mojibake_reparados),
    }

    # Lista de correcciones de ubicación hechas (para el reporte)
    correcciones_ubicacion = []

    # ───── Id empleado ─────
    if "Id empleado" in df.columns:
        antes = df["Id empleado"].copy()
        df["Id empleado"] = df["Id empleado"].apply(validar_corregir_id)
        correcciones["ids_corregidos"] = (antes != df["Id empleado"]).sum()

    # ───── Centro de costo y sede (fijos) ─────
    if "Id centro de costo" in df.columns:
        df["Id centro de costo"] = "sinDefinir"
    if "Id sede donde se desempeña" in df.columns:
        df["Id sede donde se desempeña"] = "sinDefinir"

    # ───── Fechas ─────
    for campo in campos_fecha:
        if campo in df.columns:
            antes = df[campo].copy()
            df[campo] = df[campo].apply(convertir_fecha)
            correcciones["fechas_normalizadas"] += (antes.fillna("") != df[campo].fillna("")).sum()

    # ───── Dirección ─────
    campos_direccion = ["Nombre Calle", "Numero Calle", "Departamento"]
    for campo in campos_direccion:
        if campo in df.columns:
            antes = df[campo].copy()
            df[campo] = df[campo].apply(limpiar_direccion)
            correcciones["direcciones_limpiadas"] += (antes.fillna("") != df[campo].fillna("")).sum()

    # ───── Comuna: resolver nombre → código y aplicar zfill(5) ─────
    if "Comuna" in df.columns:
        antes = df["Comuna"].copy()
        resultados = df["Comuna"].apply(resolver_comuna)
        df["Comuna"] = resultados.apply(lambda r: r[0])
        correcciones["comunas_rellenadas"] = (antes.fillna("") != df["Comuna"].fillna("")).sum()

    # ───── Emails ─────
    for campo in campos_email:
        if campo in df.columns:
            antes = df[campo].copy()
            # Normalizar a minúscula
            df[campo] = df[campo].apply(convertir_email_minuscula)
            # Completar vacíos
            mask_vacios = df[campo].apply(_vacio)
            correcciones["emails_vacios_completados"] += mask_vacios.sum()
            df.loc[mask_vacios, campo] = EMAIL_DEFAULT
            # Contar los que se pasaron a minúscula (excluyendo los completados)
            cambios_case = (antes.fillna("") != df[campo].fillna("")) & ~mask_vacios
            correcciones["emails_minuscula"] += cambios_case.sum()

    # ───── Teléfonos ─────
    for campo in campos_telefono:
        if campo in df.columns:
            nuevos, estados = [], []
            for v in df[campo]:
                nuevo, estado, _ = normalizar_telefono(v)
                nuevos.append(nuevo)
                estados.append(estado)
            df[campo] = nuevos
            correcciones["telefonos_vacios_completados"] += estados.count("completado")
            correcciones["telefonos_normalizados"]       += estados.count("normalizado")

    # ───── VALIDACIONES (errores a reportar) ─────
    errores = []
    hoy = datetime.now()

    for idx, fila in df.iterrows():
        num_fila = idx + 2  # +2 porque Excel empieza en 1 y tiene header
        campos_vacios = []
        errores_fila = []

        # Campos obligatorios vacíos / validaciones por campo
        for campo in campos_obligatorios:
            if campo not in df.columns:
                continue
            valor = fila[campo]

            if _vacio(valor):
                # Los emails y teléfonos ya se completaron, no reportar como vacíos
                if campo in campos_email or campo in campos_telefono:
                    continue
                campos_vacios.append(campo)
                continue

            if campo == "Sexo" and str(valor).strip().upper() not in ["M", "F"]:
                errores_fila.append(f"Sexo (valor: '{valor}' debe ser M o F)")

            if campo == "Estado civil" and str(valor).strip().upper() not in estados_civiles_validos:
                errores_fila.append(f"Estado civil (valor: '{valor}' debe ser S, C, V, D o U)")

            if campo in campos_email and not validar_email(valor):
                errores_fila.append(f"{campo} (valor: '{valor}' no tiene formato válido)")

            if campo == "Id empleado":
                id_str = str(valor).strip()
                if len(id_str) not in [9, 10]:
                    errores_fila.append(f"Id empleado (valor: '{valor}' debe tener 9 o 10 caracteres)")
                elif not validar_rut_dv(id_str):
                    errores_fila.append(f"Id empleado (valor: '{valor}' dígito verificador incorrecto)")

        # Validar teléfonos con formato inválido (que no se pudieron normalizar)
        for campo in campos_telefono:
            if campo in df.columns:
                v = fila[campo]
                if not _vacio(v) and v != TELEFONO_DEFAULT:
                    _, _, valido = normalizar_telefono(v)
                    if not valido:
                        errores_fila.append(f"{campo} (valor: '{v}' no tiene formato válido)")

        # ───── Validaciones entre fechas ─────
        f_nac   = _parsear_fecha(fila.get("Fecha de nacimiento"))
        f_ini   = _parsear_fecha(fila.get("Fecha de inicio del contrato"))
        f_term  = _parsear_fecha(fila.get("Fecha de término del contrato"))

        if f_nac and f_nac >= hoy:
            errores_fila.append(f"Fecha de nacimiento ({fila['Fecha de nacimiento']}) no puede ser futura")

        if f_nac and f_ini and f_nac >= f_ini:
            errores_fila.append(
                f"Fecha de nacimiento ({fila['Fecha de nacimiento']}) debe ser anterior a Fecha inicio contrato ({fila['Fecha de inicio del contrato']})"
            )

        if f_ini and f_term and f_ini > f_term:
            errores_fila.append(
                f"Fecha inicio contrato ({fila['Fecha de inicio del contrato']}) no puede ser posterior a Fecha término ({fila['Fecha de término del contrato']})"
            )

        # ───── Validar / Corregir comuna-región-ciudad ─────
        if "Comuna" in df.columns:
            region_nueva, ciudad_nueva, cambios_ubic, error_ubic = corregir_ubicacion(
                fila.get("Comuna"),
                fila.get("Region") if "Region" in df.columns else None,
                fila.get("Ciudad") if "Ciudad" in df.columns else None,
            )

            if error_ubic:
                # La comuna no existe en el maestro → reportar como error
                errores_fila.append(error_ubic)
            elif cambios_ubic:
                # Aplicar las correcciones al DataFrame
                if "Region" in df.columns:
                    df.at[idx, "Region"] = region_nueva
                if "Ciudad" in df.columns:
                    df.at[idx, "Ciudad"] = ciudad_nueva

                # Registrar cambios para el reporte y las métricas
                for cambio in cambios_ubic:
                    if cambio.startswith("Región"):
                        correcciones["regiones_corregidas"] += 1
                    elif cambio.startswith("Ciudad"):
                        correcciones["ciudades_corregidas"] += 1
                correcciones_ubicacion.append({
                    "fila":    num_fila,
                    "comuna":  str(fila.get("Comuna", "")).strip(),
                    "cambios": cambios_ubic,
                })

        if campos_vacios or errores_fila:
            errores.append((num_fila, campos_vacios, errores_fila))

    # ───── Limpieza final: reemplazar valores nulos visibles por string vacío ─────
    # pandas convierte NaN/None a string "None" o "nan" cuando exportamos a Excel
    # o mostramos en streamlit. Los reemplazamos por "" para que queden en blanco.
    df = df.fillna("")
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).apply(lambda v: "" if _vacio(v) else v)

    return df, errores, total_original, filas_eliminadas, correcciones, correcciones_ubicacion


# ─────────────────────────────────────────────
#  UI
# ─────────────────────────────────────────────

# Instrucciones del formato del archivo
st.markdown("### 📄 Formato del archivo")

col_i1, col_i2, col_i3 = st.columns(3)

with col_i1:
    st.markdown(
        '<div style="background:white;border:1px solid #E8EEF3;border-radius:12px;padding:1.25rem;height:100%;">'
        '<div style="color:#1EBBEF;font-weight:700;font-size:0.75rem;letter-spacing:0.5px;margin-bottom:0.5rem;">NOMBRE SUGERIDO</div>'
        '<div style="color:#1A3A5F;font-size:1rem;font-weight:600;margin-bottom:0.25rem;">maestro_empleados</div>'
        '<div style="color:#8B9DAE;font-size:0.8rem;">Puedes subir el archivo con cualquier nombre</div>'
        '</div>',
        unsafe_allow_html=True,
    )

with col_i2:
    st.markdown(
        '<div style="background:white;border:1px solid #E8EEF3;border-radius:12px;padding:1.25rem;height:100%;">'
        '<div style="color:#1EBBEF;font-weight:700;font-size:0.75rem;letter-spacing:0.5px;margin-bottom:0.5rem;">EXTENSIÓN</div>'
        '<div style="color:#1A3A5F;font-size:1rem;font-weight:600;margin-bottom:0.25rem;">.xlsm o .xlsx</div>'
        '<div style="color:#8B9DAE;font-size:0.8rem;">Formato Excel con o sin macros habilitadas</div>'
        '</div>',
        unsafe_allow_html=True,
    )

with col_i3:
    st.markdown(
        '<div style="background:white;border:1px solid #E8EEF3;border-radius:12px;padding:1.25rem;height:100%;">'
        '<div style="color:#1EBBEF;font-weight:700;font-size:0.75rem;letter-spacing:0.5px;margin-bottom:0.5rem;">HOJA</div>'
        '<div style="color:#1A3A5F;font-size:1rem;font-weight:600;margin-bottom:0.25rem;">Empleados</div>'
        '<div style="color:#8B9DAE;font-size:0.8rem;">El <strong>encabezado</strong> debe estar en la <strong>línea 1</strong></div>'
        '</div>',
        unsafe_allow_html=True,
    )

st.markdown("")  # espaciado

st.markdown("### 📤 Subir archivo")
archivo = st.file_uploader(
    "Selecciona el archivo Excel",
    type=["xlsm", "xlsx"],
    help="El archivo debe tener una hoja llamada 'Empleados' con el encabezado en la línea 1.",
)

if archivo:
    # ─── Validaciones previas del archivo ───
    # Validar que la hoja "Empleados" exista
    try:
        xl = pd.ExcelFile(archivo)
        hojas = xl.sheet_names
        archivo.seek(0)  # resetear el cursor para que procesar_archivo pueda leerlo
        if "Empleados" not in hojas:
            st.error(
                f"❌ El archivo no contiene una hoja llamada **'Empleados'**. "
                f"Hojas encontradas: {', '.join(hojas)}. "
                f"Asegúrate que los datos estén en una hoja llamada exactamente 'Empleados'."
            )
            st.stop()
    except Exception as e:
        st.error(f"❌ No se pudo leer el archivo Excel: {e}")
        st.stop()

    with st.spinner("Procesando archivo..."):
        try:
            df, errores, total_original, filas_eliminadas, correcciones, correcciones_ubicacion = procesar_archivo(archivo)

            # ─── Métricas generales ───
            st.markdown("### Resumen general")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Filas originales", total_original)
            c2.metric("Filas eliminadas (no A)", filas_eliminadas)
            c3.metric("Filas procesadas", len(df))
            c4.metric("Filas con errores", len(errores))

            # ─── Resumen de correcciones aplicadas ───
            st.markdown("### Correcciones aplicadas automáticamente")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("📅 Fechas normalizadas",        correcciones["fechas_normalizadas"])
            c2.metric("🆔 IDs corregidos",             correcciones["ids_corregidos"])
            c3.metric("📮 Comunas con ceros",          correcciones["comunas_rellenadas"])
            c4.metric("🏠 Direcciones limpiadas",      correcciones["direcciones_limpiadas"])

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("✉️ Emails a minúsculas",         correcciones["emails_minuscula"])
            c2.metric("✉️ Emails completados",         correcciones["emails_vacios_completados"])
            c3.metric("📞 Teléfonos normalizados",     correcciones["telefonos_normalizados"])
            c4.metric("📞 Teléfonos completados",      correcciones["telefonos_vacios_completados"])

            c1, c2 = st.columns(2)
            c1.metric("🗺️ Regiones corregidas",         correcciones["regiones_corregidas"])
            c2.metric("🏙️ Ciudades corregidas",         correcciones["ciudades_corregidas"])

            if correcciones["caracteres_reparados"] > 0:
                st.info(
                    f"🔤 Se repararon **{correcciones['caracteres_reparados']}** celda(s) "
                    f"con caracteres corruptos (ej: 'AVENDA√ëO' → 'AVENDAÑO')."
                )

            # Panel de correcciones de ubicación (detalle)
            if correcciones_ubicacion:
                with st.expander(
                    f"Ver detalle de correcciones de ubicación ({len(correcciones_ubicacion)} fila(s))",
                    expanded=False,
                ):
                    for item in correcciones_ubicacion:
                        st.markdown(f"**Fila {item['fila']}** — Comuna `{item['comuna']}`:")
                        for c in item["cambios"]:
                            st.markdown(f"- {c}")

            # ─── Estado de validación ───
            if errores:
                st.warning(f"⚠️ Se encontraron errores en {len(errores)} fila(s).")
                with st.expander("Ver detalle de errores", expanded=False):
                    reporte = ""
                    for fila, vacios, errs in errores:
                        reporte += f"**Fila {fila}:**\n"
                        for campo in vacios:
                            reporte += f"- `{campo}` está vacío\n"
                        for error in errs:
                            reporte += f"- {error}\n"
                        reporte += "\n"
                    st.markdown(reporte)
            else:
                st.success("✅ ¡Todo correcto! No hay errores en el archivo.")

            # ─── Vista previa de la tabla ───
            st.markdown("### Vista previa del archivo corregido")
            n_preview = st.slider("Número de filas a mostrar", 5, min(100, len(df)) if len(df) > 5 else 5, min(10, len(df)))

            # Selector para ver solo filas con errores
            solo_errores = False
            if errores:
                solo_errores = st.checkbox("Ver solo filas con errores", value=False)

            if solo_errores:
                filas_error = [f - 2 for f, _, _ in errores]
                df_preview = df.iloc[filas_error].head(n_preview)
            else:
                df_preview = df.head(n_preview)

            st.dataframe(df_preview, use_container_width=True, hide_index=False)

            # ─── Reporte en texto ───
            reporte_txt = "REPORTE DE VALIDACION\n" + "=" * 60 + "\n\n"
            reporte_txt += f"Filas originales: {total_original}\n"
            reporte_txt += f"Filas eliminadas (no A): {filas_eliminadas}\n"
            reporte_txt += f"Filas procesadas: {len(df)}\n\n"
            reporte_txt += "CORRECCIONES APLICADAS\n" + "-" * 40 + "\n"
            for k, v in correcciones.items():
                reporte_txt += f"  {k.replace('_', ' ').capitalize()}: {v}\n"
            reporte_txt += "\n"

            if correcciones_ubicacion:
                reporte_txt += f"DETALLE DE CORRECCIONES DE UBICACIÓN ({len(correcciones_ubicacion)} fila(s))\n"
                reporte_txt += "-" * 40 + "\n"
                for item in correcciones_ubicacion:
                    reporte_txt += f"Fila {item['fila']} - Comuna {item['comuna']}:\n"
                    for c in item["cambios"]:
                        reporte_txt += f"  - {c}\n"
                reporte_txt += "\n"

            if errores:
                reporte_txt += f"ERRORES ENCONTRADOS - {len(errores)} fila(s)\n" + "-" * 40 + "\n"
                for fila, vacios, errs in errores:
                    reporte_txt += f"Fila {fila}:\n"
                    for campo in vacios:
                        reporte_txt += f"  - '{campo}' está vacío\n"
                    for error in errs:
                        reporte_txt += f"  - {error}\n"
                    reporte_txt += "\n"
            else:
                reporte_txt += "Sin errores.\n"

            # ─── Descargas ───
            buffer = io.BytesIO()
            df.to_excel(buffer, index=False)
            buffer.seek(0)

            st.markdown("---")
            st.markdown("### 📥 Descargar resultados")
            col_a, col_b = st.columns(2)
            with col_a:
                st.download_button(
                    label="📥 Descargar Excel corregido",
                    data=buffer,
                    file_name="importacion_corregido.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            with col_b:
                st.download_button(
                    label="📄 Descargar reporte de errores",
                    data=reporte_txt.encode("utf-8"),
                    file_name="reporte_errores.txt",
                    mime="text/plain",
                )

        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")
            import traceback
            with st.expander("Ver detalle técnico"):
                st.code(traceback.format_exc())

# Footer
aplicar_footer()
