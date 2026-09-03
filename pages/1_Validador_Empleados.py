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

# Mapeo de textos a códigos de moneda cotización
MONEDA_COTIZACION_MAPEO = {
    "U": "U", "UF": "U", "UNIDAD DE FOMENTO": "U",
    "P": "P", "PESOS": "P", "PESO": "P", "CLP": "P",
    "%": "%", "7%": "%", "7 %": "%",
    "F": "F", "7% + UF": "F", "7% +UF": "F", "7%+UF": "F", "7 % + UF": "F",
    "Z": "Z", "7% + UF + PESOS": "Z", "7%+UF+PESOS": "Z", "7% + UF + PESO": "Z",
}

# Mapeo de textos a códigos de tipo de contrato
TIPO_CONTRATO_MAPEO = {
    "F": "F", "PLAZO FIJO": "F", "A PLAZO FIJO": "F", "FIJO": "F",
    "I": "I", "INDEFINIDO": "I", "A INDEFINIDO": "I",
    "O": "O", "POR OBRA": "O", "POR OBRA O FAENA": "O", "OBRA O FAENA": "O", "OBRA": "O", "FAENA": "O",
    "E": "E", "APRENDIZAJE": "E", "DE APRENDIZAJE": "E", "CONTRATO DE APRENDIZAJE": "E",
    "H": "H", "HONORARIOS": "H",
}

# Mapeo de textos a códigos de estado civil
ESTADO_CIVIL_MAPEO = {
    "S": "S", "SOLTERO": "S", "SOLTERA": "S", "SOLTERO/A": "S",
    "C": "C", "CASADO": "C", "CASADA": "C", "CASADO/A": "C",
    "V": "V", "VIUDO": "V", "VIUDA": "V", "VIUDO/A": "V",
    "D": "D", "DIVORCIADO": "D", "DIVORCIADA": "D", "DIVORCIADO/A": "D",
    "U": "U", "CONVIVIENTE": "U", "CONVIVIENTE CIVIL": "U",
    "UNION CIVIL": "U", "UNION CIVIL": "U",
}

formatos_posibles = [
    "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d",
    "%Y/%m/%d", "%d-%m-%y", "%d/%m/%y", "%m/%d/%Y",
]

TELEFONO_DEFAULT = "+56922222222"
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


def limpiar_nombre_calle(valor):
    """Limpia el Nombre Calle: elimina caracteres especiales, números
    y palabras como depto/departamento/bloque/block y todo lo que las sigue."""
    if _vacio(valor):
        return valor
    texto = str(valor).strip()
    # 1. Quitar caracteres especiales (mantener letras, números y espacios por ahora)
    texto = re.sub(r"[^a-zA-ZáéíóúÁÉÍÓÚñÑ0-9 ]", " ", texto)
    # 2. Cortar desde palabras tipo depto/bloque en adelante
    texto = re.split(r"\b(?:depto|dpto|depto\.?|departamento|bloque|block|blok)\b",
                     texto, flags=re.IGNORECASE)[0]
    # 3. Eliminar todos los números
    texto = re.sub(r"\d+", "", texto)
    # 4. Eliminar "N" suelta que queda de "N°" / "Nº"
    texto = re.sub(r"\b[Nn]\b", "", texto)
    # 5. Normalizar espacios múltiples
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def convertir_email_minuscula(valor):
    if _vacio(valor):
        return valor
    return str(valor).strip().lower()


# ─────────────────────────────────────────────
#  MAESTROS DE AFP Y SALUD
# ─────────────────────────────────────────────

AFP_MAESTRO = {
    "capital":   "AFP Capital",
    "cuprum":    "AFP Cuprum",
    "habitat":   "AFP Habitat",
    "modelo":    "AFP Modelo",
    "planvital": "AFP Planvital",
    "provida":   "AFP Provida",
    "uno":       "AFP UNO",
    "canaempu":  "Canaemput",
    "capremer":  "Capremer",
    "empart":    "Empart",
    "sss":       "Servicio Seguro Social",
    "afp":       "Sin definir",
    "triomar":   "Triomar",
}

SALUD_MAESTRO = {
    "fonasa":       "FONASA",
    "isalud":       "ISALUD ISAPRE de Codelco LTDA",
    "bancoestado":  "ISAPRE Banco Estado",
    "banmedica":    "ISAPRE Banmedica",
    "chuquicamata": "ISAPRE Chuquicamata",
    "colmena":      "ISAPRE Colmena",
    "consalud":     "ISAPRE Consalud",
    "cruzblanca":   "ISAPRE Cruz-Blanca",
    "cruzdelnorte": "ISAPRE Cruz del Norte",
    "esencial":     "ISAPRE Esencial",
    "fusat":        "ISAPRE Fusat",
    "nuevamasvida": "ISAPRE Nueva Mas Vida",
    "vidatres":     "ISAPRE Vidatres",
    "isapre":       "Sin definir",
}


def _normalizar_aseguradora(texto):
    """Quita prefijos tipo 'AFP', 'ISAPRE', 'Sociedad', etc. y espacios/tildes."""
    if _vacio(texto):
        return ""
    s = _normalizar_texto(texto)
    # Quitar prefijos comunes
    prefijos = ["afp ", "isapre ", "isapre de ", "sociedad ", "s.a.", "sa ", "ltda", "ltda."]
    for pref in prefijos:
        if s.startswith(pref):
            s = s[len(pref):]
    # Quitar sufijos comunes
    for suf in [" s.a.", " sa", " ltda.", " ltda", " spa"]:
        if s.endswith(suf):
            s = s[:-len(suf)]
    # Quitar espacios, guiones y caracteres raros
    return s.replace(" ", "").replace("-", "").replace(".", "").strip()


def resolver_afp(valor):
    """Convierte cualquier variante de nombre de AFP al ID oficial.

    Ej: 'AFP Capital' → 'capital', 'Habitat S.A.' → 'habitat'.
    Retorna (id_resuelto, cambio_realizado) donde cambio puede ser None.
    """
    if _vacio(valor):
        return valor, None

    original = str(valor).strip()
    normalizado = _normalizar_aseguradora(original)

    # Si ya es el ID exacto (minúscula) → no hay cambio
    if original.lower() in AFP_MAESTRO:
        return original.lower(), None

    # Buscar por nombre normalizado
    for afp_id in AFP_MAESTRO:
        if normalizado == afp_id:
            if original != afp_id:
                return afp_id, f"AFP convertida a ID: '{original}' → '{afp_id}'"
            return afp_id, None

    # Buscar por nombre completo normalizado del maestro
    for afp_id, nombre_oficial in AFP_MAESTRO.items():
        if _normalizar_aseguradora(nombre_oficial) == normalizado:
            return afp_id, f"AFP convertida a ID: '{original}' → '{afp_id}'"

    # No se pudo resolver — devolver sin cambio (se marcará como error después)
    return original, None


def resolver_salud(valor):
    """Convierte cualquier variante de nombre de Isapre/Fonasa al ID oficial.

    Ej: 'ISAPRE Banmedica' → 'banmedica', 'Cruz Blanca' → 'cruzblanca'.
    Retorna (id_resuelto, cambio_realizado) donde cambio puede ser None.
    """
    if _vacio(valor):
        return valor, None

    original = str(valor).strip()
    normalizado = _normalizar_aseguradora(original)

    # Si ya es el ID exacto (minúscula) → no hay cambio
    if original.lower() in SALUD_MAESTRO:
        return original.lower(), None

    # Buscar por nombre normalizado
    for salud_id in SALUD_MAESTRO:
        if normalizado == salud_id:
            if original != salud_id:
                return salud_id, f"Salud convertida a ID: '{original}' → '{salud_id}'"
            return salud_id, None

    # Buscar por nombre completo normalizado del maestro
    for salud_id, nombre_oficial in SALUD_MAESTRO.items():
        if _normalizar_aseguradora(nombre_oficial) == normalizado:
            return salud_id, f"Salud convertida a ID: '{original}' → '{salud_id}'"

    # No se pudo resolver — devolver sin cambio (se marcará como error después)
    return original, None


def normalizar_telefono(valor):
    """Normaliza a formato '+56XXXXXXXXX' (todo pegado con + al inicio).
    Retorna (valor_normalizado, cambio, valido)."""
    if _vacio(valor):
        return TELEFONO_DEFAULT, "completado", True

    original = str(valor).strip()
    # Extraer solo dígitos
    digitos = re.sub(r"\D", "", original)

    # Casos posibles
    if digitos.startswith("56") and len(digitos) == 11:
        # Ya trae código país + 9 dígitos → +56XXXXXXXXX
        formateado = f"+{digitos}"
    elif len(digitos) == 9:
        # 9 dígitos sin código país → agregar 56 adelante
        formateado = f"+56{digitos}"
    elif len(digitos) == 8:
        # Fijo antiguo sin código área — asumimos Santiago (2)
        formateado = f"+562{digitos}"
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
#  REGLAS DE NORMALIZACIÓN (campos de contrato / previsión)
# ─────────────────────────────────────────────

# Gentilicio → nombre de país. Las claves se normalizan (minúsculas, sin tildes)
# al construir el índice, así que basta con agregar la fila nueva aquí.
NACIONALIDAD_MAPEO = {
    "peruana":     "Perú",
    "peruano":     "Perú",
    "uruguaya":    "Uruguay",
    "uruguayo":    "Uruguay",
    "chilena":     "Chile",
    "chileno":     "Chile",
    "colombiana":  "Colombia",
    "colombiano":  "Colombia",
    "venezolana":  "Venezuela",
    "venezolano":  "Venezuela",
}
_INDICE_NACIONALIDAD = {_normalizar_texto(k): v for k, v in NACIONALIDAD_MAPEO.items()}

NACION_POR_DEFECTO = "Chile"

# Bancos que llegan con nombre comercial o de la cooperativa y hay que mapear.
BANCO_MAPEO = {
    "NOVA":                                 "BCI",
    "COOPERATIVA PERSONAL U. DE CHILE LTDA": "COOPEUCH",
}
_INDICE_BANCO = {_normalizar_texto(k): v for k, v in BANCO_MAPEO.items()}
BANCO_SIN_DEFINIR = "NOBANCO"

# Columnas que tocan las 14 reglas de normalización, con sus alias aceptados.
# Se usa para el diagnóstico agregado: ocupación y columnas ausentes.
COLUMNAS_REGLAS = [
    ("1",     ("Apellido materno",)),
    ("2",     ("Id nación", "Id nacion")),
    ("3",     ("¿Es expatriado?",)),
    ("4",     ("Estado de jubilación",)),
    ("5",     ("Sistema de pensiones",)),
    ("6",     ("Id banco", "Banco")),
    ("7",     ("Monto cotizado en la Isapre", "Monto cotizado en Isapre")),
    ("8",     ("Monto cotizado en la Isapre en UF", "Monto cotizado en Isapre en UF")),
    ("10",    ("¿Jornada parcial?",)),
    ("10-hrs",("Horas de trabajo semanales",)),
    ("11",    ("Fecha de inicio de vacaciones",)),
    ("12",    ("Fecha de incorporación al seguro de cesantía",)),
    ("13",    ("¿Cotiza seguro de cesantía?",)),
    ("14",    ("Modalidad del contrato",)),
    ("ref",   ("Fecha de inicio del contrato",)),
    ("salud", ("ID INSTITUCION DE SALUD", "ID INSTITUCIÓN DE SALUD")),
]

# Umbral de jornada parcial: <= 30 horas semanales es jornada parcial.
JORNADA_PARCIAL_MAX_HORAS = 30

# Log crudo del monto Isapre: muestreo por celda del valor tal como viene, con
# fila e Id empleado, más el volcado a consola. Es solo para depurar el parseo
# contra archivos reales; queda en False porque recolecta RUTs y ensucia el log.
# NO afecta a los paneles de Verificación 7-8, Detalle de reglas ni Diagnóstico,
# que se muestran siempre.
DEBUG_MONTO_ISAPRE = False


def _solo_ceros(valor) -> bool:
    """True si el valor son puros ceros: '0', '00', '0.0', '0,00', '000-0'."""
    if _vacio(valor):
        return False
    s = re.sub(r"[.,\-\s]", "", str(valor).strip())
    return s != "" and set(s) == {"0"}


def _buscar_columna(df, *candidatos):
    """Devuelve el nombre real de la primera columna que calce (sin tildes/case)."""
    normalizadas = {_normalizar_texto(c): c for c in df.columns}
    for cand in candidatos:
        real = normalizadas.get(_normalizar_texto(cand))
        if real is not None:
            return real
    return None


def _canon(valor) -> str:
    """Forma canónica para comparar 'antes vs después' sin falsos positivos.

    NaN, None, '' y '  ' colapsan a ''. Los números se comparan por su texto
    normalizado, para que un cambio de tipo ('0' → 0) no cuente como cambio.
    """
    if _vacio(valor):
        return ""
    s = str(valor).strip()
    try:
        f = float(s.replace(",", "."))
        return str(int(f)) if f == int(f) else str(f)
    except (ValueError, TypeError):
        return s


def _medir_columna(antes, despues) -> dict:
    """Compara dos versiones de una columna y mide qué cambió DE VERDAD."""
    a = antes.apply(_canon)
    d = despues.apply(_canon)
    cambiadas = (a != d)
    return {
        "filas":            int(len(a)),
        "no_vacias_antes":  int((a != "").sum()),
        "vacias_antes":     int((a == "").sum()),
        "cambios_reales":   int(cambiadas.sum()),
        "rellenos":         int((cambiadas & (a == "")).sum()),
        "sobrescrituras":   int((cambiadas & (a != "")).sum()),
    }


def _completar_vacios(df, col, valor) -> int:
    """Rellena las celdas vacías de `col` con `valor`. Devuelve cuántas cambió."""
    df[col] = df[col].astype("object")
    mask = df[col].apply(_vacio).astype(bool)
    df.loc[mask, col] = valor
    return int(mask.sum())


# ───── Regla 1: Apellido materno ─────
def limpiar_apellido_materno(valor):
    """'.' o ceros ('0', '00', ...) → campo vacío."""
    if _vacio(valor):
        return valor
    s = str(valor).strip()
    if s == "." or _solo_ceros(s):
        return ""
    return valor


# ───── Regla 2: Id nación ─────
def normalizar_nacionalidad(valor):
    """Gentilicio → país (sin tildes, case-insensitive). Vacío → 'Chile'."""
    if _vacio(valor):
        return NACION_POR_DEFECTO
    pais = _INDICE_NACIONALIDAD.get(_normalizar_texto(valor))
    return pais if pais else str(valor).strip()


# ───── Regla 3: ¿Es expatriado? ─────
def calcular_expatriado(id_nacion):
    """Se evalúa DESPUÉS de normalizar la nación: Chile → 'N', el resto → 'T'."""
    return "N" if _normalizar_texto(id_nacion) == _normalizar_texto(NACION_POR_DEFECTO) else "T"


# ───── Regla 6: Banco ─────
def normalizar_banco(valor):
    """Mapea bancos conocidos; vacío o en ceros → 'NOBANCO'."""
    if _vacio(valor) or _solo_ceros(valor):
        return BANCO_SIN_DEFINIR
    return _INDICE_BANCO.get(_normalizar_texto(valor), str(valor).strip())


# ───── Reglas 7 y 8: Institución de salud ─────
def es_fonasa(valor) -> bool:
    """True si la institución de salud es Fonasa (ignora case, tildes y prefijos)."""
    return _normalizar_aseguradora(valor) == "fonasa"


# ───── Regla 9: Monto cotizado Isapre ─────
def parsear_monto(valor):
    """Convierte el monto a número aceptando coma decimal.

    El cliente manda el monto con coma ('4,5') y en el Excel en formato General
    queda como texto/entero, perdiendo el decimal. Se toma el valor crudo de la
    celda (antes de cualquier cast) y se fuerza el parseo cambiando ',' por '.'.
    Devuelve (valor_convertido, hubo_cambio).
    """
    if _vacio(valor):
        return valor, False

    crudo = str(valor).strip()
    limpio = re.sub(r"[^\d,.\-]", "", crudo)

    # Si trae ambos separadores, el último es el decimal y el otro es de miles.
    if "," in limpio and "." in limpio:
        if limpio.rfind(",") > limpio.rfind("."):
            limpio = limpio.replace(".", "").replace(",", ".")
        else:
            limpio = limpio.replace(",", "")
    elif "," in limpio:
        limpio = limpio.replace(",", ".")

    try:
        numero = float(limpio)
    except ValueError:
        return valor, False

    convertido = int(numero) if numero == int(numero) else numero
    return convertido, str(convertido) != crudo


def clasificar_formato_monto(valor) -> str:
    """Clasifica el formato CRUDO de la celda, antes de cualquier cast.

    - 'coma'        → '4,5'   el decimal llegó con coma (se puede parsear)
    - 'punto'       → '4.5'   el decimal llegó con punto (se puede parsear)
    - 'sin_decimal' → '45'    NO hay decimal en la celda: se perdió al guardar
                              el Excel, no hay fix posible desde el código
    - 'no_numerico' → '%', 'N/A', texto libre
    """
    if _vacio(valor):
        return "vacio"
    crudo = str(valor).strip()
    if not re.search(r"\d", crudo):
        return "no_numerico"
    if "," in crudo:
        return "coma"
    if "." in crudo:
        return "punto"
    return "sin_decimal"


# ───── Regla 10: ¿Jornada parcial? ─────
def jornada_parcial_por_horas(horas):
    """<= JORNADA_PARCIAL_MAX_HORAS → 'S'; > → 'N'. Sin horas → None (no se toca)."""
    if _vacio(horas):
        return None
    numero, _ = parsear_monto(horas)
    if not isinstance(numero, (int, float)):
        return None
    return "S" if float(numero) <= JORNADA_PARCIAL_MAX_HORAS else "N"


# ─────────────────────────────────────────────
#  PROCESAMIENTO PRINCIPAL
# ─────────────────────────────────────────────

def procesar_archivo(uploaded_file):
    # Detectar automáticamente la fila de encabezados (algunos archivos
    # traen una nota tipo "* Los campos obligatorios..." antes del header)
    raw = pd.read_excel(uploaded_file, sheet_name="Empleados", dtype=str, header=None)
    fila_header = 0
    for i in range(min(10, len(raw))):
        if raw.iloc[i].astype(str).str.strip().eq("Id empleado").any():
            fila_header = i
            break
    df = raw.iloc[fila_header + 1:].reset_index(drop=True)
    df.columns = [str(c).strip() if c is not None else "" for c in raw.iloc[fila_header]]
    total_original = len(df)

    # ───── Reparar caracteres corruptos (mojibake) en headers y celdas ─────
    # Ej: "AVENDA√ëO" → "AVENDAÑO", "tel√©fono" → "teléfono"
    mojibake_reparados = 0
    df.columns = [reparar_mojibake(c) for c in df.columns]
    for col in df.columns:
        # pandas >= 3: read_excel(dtype=str) produce dtype "str" (no "object"),
        # por lo que el chequeo antiguo `== object` saltaba todas las columnas
        if df[col].dtype == object or pd.api.types.is_string_dtype(df[col]):
            antes = df[col].copy()
            df[col] = df[col].apply(reparar_mojibake)
            mojibake_reparados += int((antes.fillna("") != df[col].fillna("")).sum())

    # Filtrar solo empleados con Situación A o F
    if "Situación" in df.columns:
        df = df[df["Situación"].str.strip().str.upper().isin(["A", "F"])].reset_index(drop=True)
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
        "afp_normalizadas":          0,
        "salud_normalizadas":        0,
        "estados_civiles_normalizados": 0,
        "tipos_contrato_normalizados": 0,
        "monedas_normalizadas": 0,
        "valores_defecto_completados": 0,
        "reglas_normalizacion":       0,
    }

    # Detalle por regla nueva (para separarlo de los defaults que ya existían)
    detalle_reglas = {}

    def _anotar_regla(nombre, columna):
        """Registra qué columna toca cada regla; el conteo se mide al final."""
        columnas_por_regla.setdefault(nombre, columna)

    # Lista de correcciones de ubicación hechas (para el reporte)
    correcciones_ubicacion = []

    # ───── Id empleado ─────
    if "Id empleado" in df.columns:
        antes = df["Id empleado"].copy()
        df["Id empleado"] = df["Id empleado"].apply(validar_corregir_id)
        correcciones["ids_corregidos"] = int((antes.fillna("") != df["Id empleado"].fillna("")).sum())

    # ───── Centro de costo y sede (fijos) ─────
    if "Id centro de costo" in df.columns:
        df["Id centro de costo"] = "sinDefinir"
    if "Id sede donde se desempeña" in df.columns:
        df["Id sede donde se desempeña"] = "sinDefinir"

    # ───── Valores por defecto para campos vacíos ─────
    DEFAULTS_CAMPOS = {
        "Licencia de conducir": "N",
        "Profesión":            "sinDefinir",
        "Nivel de estudio":     "0",
    }
    for campo, valor_defecto in DEFAULTS_CAMPOS.items():
        if campo in df.columns:
            # .astype("object") evita que una columna 100% vacía (float64 en
            # pandas >= 3) rechace la asignación de strings
            df[campo] = df[campo].astype("object")
            mask_def = df[campo].apply(_vacio).astype(bool)
            df.loc[mask_def, campo] = valor_defecto
            correcciones["valores_defecto_completados"] += int(mask_def.sum())

    # ───── Fechas ─────
    for campo in campos_fecha:
        if campo in df.columns:
            antes = df[campo].copy()
            df[campo] = df[campo].apply(convertir_fecha)
            correcciones["fechas_normalizadas"] += int((antes.fillna("") != df[campo].fillna("")).sum())

    # ───── Dirección ─────
    # Nombre Calle: limpieza especial (sin números, sin depto/bloque)
    if "Nombre Calle" in df.columns:
        antes = df["Nombre Calle"].copy()
        df["Nombre Calle"] = df["Nombre Calle"].apply(limpiar_nombre_calle)
        correcciones["direcciones_limpiadas"] += int((antes.fillna("") != df["Nombre Calle"].fillna("")).sum())

    # Numero Calle y Departamento: limpieza estándar
    campos_direccion = ["Numero Calle", "Departamento"]
    for campo in campos_direccion:
        if campo in df.columns:
            antes = df[campo].copy()
            df[campo] = df[campo].apply(limpiar_direccion)
            correcciones["direcciones_limpiadas"] += int((antes.fillna("") != df[campo].fillna("")).sum())

    # ───── Comuna: resolver nombre → código y aplicar zfill(5) ─────
    if "Comuna" in df.columns:
        antes = df["Comuna"].copy()
        resultados = df["Comuna"].apply(resolver_comuna)
        df["Comuna"] = resultados.apply(lambda r: r[0])
        correcciones["comunas_rellenadas"] = int((antes.fillna("") != df["Comuna"].fillna("")).sum())

    # ───── Moneda cotización: normalizar texto a código ─────
    if "Moneda de la cotización" in df.columns:
        antes = df["Moneda de la cotización"].copy()
        def _normalizar_moneda_cotizacion(valor):
            if pd.isna(valor) or str(valor).strip() == "":
                return valor
            clave = str(valor).strip().upper()
            return MONEDA_COTIZACION_MAPEO.get(clave, valor)
        df["Moneda de la cotización"] = df["Moneda de la cotización"].apply(_normalizar_moneda_cotizacion)
        correcciones["monedas_normalizadas"] = int((antes.fillna("") != df["Moneda de la cotización"].fillna("")).sum())

    # ───── Tipo de contrato: normalizar texto a código ─────
    if "Tipo del contrato" in df.columns:
        antes = df["Tipo del contrato"].copy()
        def _normalizar_tipo_contrato(valor):
            if pd.isna(valor) or str(valor).strip() == "":
                return valor
            clave = str(valor).strip().upper()
            return TIPO_CONTRATO_MAPEO.get(clave, valor)
        df["Tipo del contrato"] = df["Tipo del contrato"].apply(_normalizar_tipo_contrato)
        correcciones["tipos_contrato_normalizados"] = int((antes.fillna("") != df["Tipo del contrato"].fillna("")).sum())

    # ───── Estado civil: normalizar texto a código ─────
    if "Estado civil" in df.columns:
        antes = df["Estado civil"].copy()
        def _normalizar_estado_civil(valor):
            if pd.isna(valor) or str(valor).strip() == "":
                return valor
            clave = str(valor).strip().upper()
            return ESTADO_CIVIL_MAPEO.get(clave, valor)
        df["Estado civil"] = df["Estado civil"].apply(_normalizar_estado_civil)
        correcciones["estados_civiles_normalizados"] = int((antes.fillna("") != df["Estado civil"].fillna("")).sum())

    # ───── AFP: normalizar a ID oficial ─────
    if "Id AFP" in df.columns:
        antes = df["Id AFP"].copy()
        resultados = df["Id AFP"].apply(resolver_afp)
        df["Id AFP"] = resultados.apply(lambda r: r[0])
        correcciones["afp_normalizadas"] = int((antes.fillna("") != df["Id AFP"].fillna("")).sum())

    # ───── Salud: normalizar a ID oficial ─────
    col_salud = next((c for c in df.columns if c.strip().upper().replace("Ó","O").replace("Ú","U") 
                      in ["ID INSTITUCION DE SALUD", "ID INSTITUCIÓN DE SALUD"]), None)
    if col_salud:
        antes = df[col_salud].copy()
        resultados = df[col_salud].apply(resolver_salud)
        df[col_salud] = resultados.apply(lambda r: r[0])
        # Convertir a minúscula
        df[col_salud] = df[col_salud].apply(
            lambda v: str(v).strip().lower() if pd.notna(v) and str(v).strip() != "" else v
        )
        correcciones["salud_normalizadas"] = int((antes.fillna("") != df[col_salud].fillna("")).sum())

    # ───── Reglas de normalización de contrato / previsión ─────
    # Se aplican en el orden pedido: la nación se normaliza antes de calcular
    # expatriado, y el monto Isapre se parsea antes de los overrides de Fonasa.
    # Resumen del formato crudo de los montos Isapre (ver DEBUG_MONTO_ISAPRE)
    log_monto_isapre = {}

    # Snapshot ANTES de aplicar las reglas: el detalle se mide comparando
    # contra esta copia, no acumulando contadores por fila procesada.
    df_antes_reglas = df.copy()
    columnas_por_regla = {}

    # 1) Apellido materno: '.' o ceros → vacío
    col = _buscar_columna(df, "Apellido materno")
    if col:
        antes = df[col].copy()
        df[col] = df[col].astype("object").apply(limpiar_apellido_materno)
        _anotar_regla("1. Apellido materno ('.'/ceros → vacío)", col)

    # 2) Id nación: gentilicio → país, y vacío → Chile
    col_nacion = _buscar_columna(df, "Id nación", "Id nacion")
    if col_nacion:
        antes = df[col_nacion].copy()
        df[col_nacion] = df[col_nacion].astype("object").apply(normalizar_nacionalidad)
        _anotar_regla("2. Id nación (gentilicio → país / vacío → Chile)", col_nacion)

    # 3) ¿Es expatriado? — se calcula con la nación ya normalizada
    col_exp = _buscar_columna(df, "¿Es expatriado?")
    if col_exp and col_nacion:
        antes = df[col_exp].copy()
        df[col_exp] = df[col_nacion].apply(calcular_expatriado)
        _anotar_regla("3. ¿Es expatriado? (derivado de Id nación)", col_exp)

    # 4-5, 13-14) Vacíos con valor por defecto
    DEFAULTS_REGLAS = [
        ("4",  ("Estado de jubilación",),        "0"),
        ("5",  ("Sistema de pensiones",),        "N"),
        ("13", ("¿Cotiza seguro de cesantía?",), "S"),
        ("14", ("Modalidad del contrato",),      "C"),
    ]
    for num, nombres, valor_defecto in DEFAULTS_REGLAS:
        col = _buscar_columna(df, *nombres)
        if col:
            _completar_vacios(df, col, valor_defecto)
            _anotar_regla(f"{num}. {nombres[0]} (vacío → '{valor_defecto}')", col)

    # 6) Banco: mapeo de los que faltan identificar
    col = _buscar_columna(df, "Id banco", "Banco")
    if col:
        antes = df[col].copy()
        df[col] = df[col].astype("object").apply(normalizar_banco)
        _anotar_regla("6. Banco (mapeo / vacío-ceros → NOBANCO)", col)

    # 9) Monto Isapre: parseo forzando la coma decimal (antes de los overrides)
    col_monto_pesos = _buscar_columna(df, "Monto cotizado en la Isapre", "Monto cotizado en Isapre")
    col_monto_uf    = _buscar_columna(df, "Monto cotizado en la Isapre en UF", "Monto cotizado en Isapre en UF")
    for col in [c for c in (col_monto_pesos, col_monto_uf) if c]:
        nuevos, parseados = [], 0
        for idx, crudo in df[col].items():
            convertido, cambio = parsear_monto(crudo)
            nuevos.append(convertido)
            if cambio:
                parseados += 1
            if DEBUG_MONTO_ISAPRE and not _vacio(crudo):
                formato = clasificar_formato_monto(crudo)
                resumen = log_monto_isapre.setdefault(col, {})
                caso = resumen.setdefault(formato, {"total": 0, "ejemplos": []})
                caso["total"] += 1
                if len(caso["ejemplos"]) < 5:
                    # Guardar fila, empleado y salud para poder rastrear el origen
                    caso["ejemplos"].append({
                        "fila":       int(idx) + fila_header + 2,
                        "empleado":   str(df.at[idx, "Id empleado"]).strip() if "Id empleado" in df.columns else "",
                        "salud":      str(df.at[idx, col_salud]).strip() if col_salud else "",
                        "crudo":      repr(crudo),
                        "convertido": repr(convertido),
                    })
        df[col] = pd.Series(nuevos, index=df.index, dtype="object")
        _anotar_regla(f"9. {col} (parseo coma decimal)", col)

    # 7-8) Si la institución de salud es Fonasa: monto en pesos → 0, monto UF → '%'
    verificacion_fonasa = {}
    if col_salud:
        mask_fonasa = df[col_salud].apply(es_fonasa).astype(bool)
        verificacion_fonasa["filas_fonasa"] = int(mask_fonasa.sum())
        for col, valor_fonasa, num in ((col_monto_pesos, 0, "7"), (col_monto_uf, "%", "8")):
            if col:
                df[col] = df[col].astype("object")
                distintos = mask_fonasa & (df[col] != valor_fonasa)
                df.loc[mask_fonasa, col] = valor_fonasa
                _anotar_regla(f"{num}. {col} (Fonasa → {valor_fonasa!r})", col)
                # Verificación post-override: qué quedó en la columna para Fonasa
                if verificacion_fonasa["filas_fonasa"]:
                    conteo = df.loc[mask_fonasa, col].astype(str).value_counts()
                    verificacion_fonasa[col] = {str(k): int(v) for k, v in conteo.items()}

    # 10) ¿Jornada parcial? vacío → según horas de trabajo semanales
    col_jornada = _buscar_columna(df, "¿Jornada parcial?")
    col_horas   = _buscar_columna(df, "Horas de trabajo semanales")
    if col_jornada and col_horas:
        _anotar_regla("10. ¿Jornada parcial? (según horas semanales)", col_jornada)
        df[col_jornada] = df[col_jornada].astype("object")
        for idx in df.index:
            if not _vacio(df.at[idx, col_jornada]):
                continue
            valor = jornada_parcial_por_horas(df.at[idx, col_horas])
            if valor is not None:
                df.at[idx, col_jornada] = valor

    # 11-12) Fechas que se completan con la fecha de inicio del contrato
    col_inicio = _buscar_columna(df, "Fecha de inicio del contrato")
    if col_inicio:
        for nombres in (("Fecha de inicio de vacaciones",),
                        ("Fecha de incorporación al seguro de cesantía",)):
            col = _buscar_columna(df, *nombres)
            if not col:
                continue
            df[col] = df[col].astype("object")
            mask = df[col].apply(_vacio).astype(bool) & ~df[col_inicio].apply(_vacio).astype(bool)
            df.loc[mask, col] = df.loc[mask, col_inicio]
            _anotar_regla(f"11-12. {col} (vacía → fecha inicio contrato)", col)

    # ── Medición real: comparar cada columna contra el snapshot previo ──
    for _nombre, _col in columnas_por_regla.items():
        if _col in df.columns and _col in df_antes_reglas.columns:
            detalle_reglas[_nombre] = _medir_columna(df_antes_reglas[_col], df[_col])
    correcciones["reglas_normalizacion"] = sum(
        m["cambios_reales"] for m in detalle_reglas.values()
    )

    # ── Diagnóstico agregado (solo conteos, sin datos personales) ──
    diagnostico = {
        "filas":            int(len(df)),
        "columnas":         int(len(df.columns)),
        "filas_originales": int(total_original),
        "filas_eliminadas": int(filas_eliminadas),
        "ocupacion":        [],
        "ausentes":         [],
    }
    for _num, _cands in COLUMNAS_REGLAS:
        _col = _buscar_columna(df, *_cands)
        if _col is None:
            diagnostico["ausentes"].append(f"{_cands[0]} (regla {_num})")
            continue
        _serie = df_antes_reglas[_col] if _col in df_antes_reglas.columns else df[_col]
        _canonica = _serie.apply(_canon)
        diagnostico["ocupacion"].append({
            "regla":      _num,
            "columna":    _col,
            "con_valor":  int((_canonica != "").sum()),
            "vacias":     int((_canonica == "").sum()),
        })

    # Paneles permanentes: se reescriben en cada carga para no arrastrar
    # resultados del archivo anterior en la misma sesión.
    st.session_state["diagnostico_maestro"] = diagnostico
    st.session_state["detalle_reglas"] = detalle_reglas
    st.session_state["verificacion_fonasa"] = verificacion_fonasa

    # Log crudo (detrás de la bandera). Se limpia siempre, si no el panel
    # quedaría mostrando los datos del archivo procesado anteriormente.
    st.session_state["debug_monto_isapre"] = (
        log_monto_isapre if (DEBUG_MONTO_ISAPRE and log_monto_isapre) else None
    )

    if DEBUG_MONTO_ISAPRE and log_monto_isapre:
        for _col, _resumen in log_monto_isapre.items():
            print(f"[DEBUG_MONTO_ISAPRE] {_col}:")
            for _formato, _caso in sorted(_resumen.items()):
                # Sin Id empleado: el log del servidor puede quedar en disco
                _ej = "; ".join(
                    f"fila {e['fila']} {e['crudo']} → {e['convertido']}"
                    for e in _caso["ejemplos"]
                )
                print(f"    {_formato:12s} {_caso['total']:6d}  ej: {_ej}")

    # ───── Emails ─────
    for campo in campos_email:
        if campo in df.columns:
            antes = df[campo].copy()
            # Normalizar a minúscula
            # .astype("object") evita que pandas re-infiera la columna como
            # float64 cuando viene 100% vacía (todos NaN), lo que rompía la
            # asignación de placeholders más abajo en pandas >= 3
            df[campo] = df[campo].apply(convertir_email_minuscula).astype("object")
            # Completar vacíos con {rut}@sincorreo.cl
            mask_vacios = df[campo].apply(_vacio).astype(bool)
            correcciones["emails_vacios_completados"] += int(mask_vacios.sum())
            if "Id empleado" in df.columns:
                df.loc[mask_vacios, campo] = (
                    df.loc[mask_vacios, "Id empleado"].astype(str).str.strip() + "@sincorreo.cl"
                )
            else:
                df.loc[mask_vacios, campo] = EMAIL_DEFAULT
            # Contar los que se pasaron a minúscula (excluyendo los completados)
            cambios_case = (antes.fillna("") != df[campo].fillna("")) & ~mask_vacios
            correcciones["emails_minuscula"] += int(cambios_case.sum())

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
        num_fila = idx + fila_header + 2  # +2 porque Excel empieza en 1 y tiene header
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
                if campo == "Tipo del contrato":
                    # No se rellena: la decisión (Indefinido vs. Plazo Fijo) es
                    # del cliente. Solo se deja como observación en el reporte.
                    errores_fila.append(
                        "OBSERVACIÓN: Tipo del contrato vacío — requiere definición del "
                        "cliente (Indefinido / Plazo Fijo). No se completa automáticamente."
                    )
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
            c2.metric("Filas eliminadas (no A/F)", filas_eliminadas)
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

            c1, c2, c3 = st.columns(3)
            c1.metric("🏦 AFP normalizadas",            correcciones["afp_normalizadas"])
            c2.metric("🏥 Salud normalizadas",          correcciones["salud_normalizadas"])
            c3.metric("⚙️ Reglas aplicadas",            correcciones["reglas_normalizacion"])

            if correcciones["valores_defecto_completados"] > 0:
                st.info(
                    f"📝 Reglas preexistentes: se completaron "
                    f"**{correcciones['valores_defecto_completados']}** celda(s) vacías en "
                    f"Licencia de conducir (N) · Profesión (sinDefinir) · Nivel de estudio (0). "
                    f"No incluye las reglas de normalización nuevas, que van aparte en "
                    f"**⚙️ Reglas aplicadas**."
                )

            if correcciones["caracteres_reparados"] > 0:
                st.info(
                    f"🔤 Se repararon **{correcciones['caracteres_reparados']}** celda(s) "
                    f"con caracteres corruptos (ej: 'AVENDA√ëO' → 'AVENDAÑO')."
                )

            # Log crudo del monto Isapre — solo con DEBUG_MONTO_ISAPRE = True
            debug_monto = st.session_state.get("debug_monto_isapre")
            if debug_monto:
                with st.expander("🔍 Formato crudo del Monto cotizado Isapre", expanded=True):
                    ETIQUETAS = {
                        "coma":        "Coma decimal ('4,5') — se parsea OK",
                        "punto":       "Punto decimal ('4.5') — se parsea OK",
                        "sin_decimal": "Sin decimal ('45') — el decimal NO viene en el Excel",
                        "no_numerico": "No numérico ('%', texto)",
                    }
                    for _col, _resumen in debug_monto.items():
                        st.markdown(f"**{_col}**")
                        st.dataframe(
                            pd.DataFrame([
                                {"Formato": ETIQUETAS.get(f, f), "Valores": caso["total"]}
                                for f, caso in sorted(_resumen.items())
                            ]),
                            use_container_width=True, hide_index=True,
                        )
                        st.caption("Ejemplos (hasta 5 por formato), con la fila del Excel de origen:")
                        st.dataframe(
                            pd.DataFrame([
                                {
                                    "Formato":     ETIQUETAS.get(f, f),
                                    "Fila Excel":  e["fila"],
                                    "Id empleado": e["empleado"],
                                    "Inst. salud": e["salud"],
                                    "Crudo":       e["crudo"],
                                    "Convertido":  e["convertido"],
                                }
                                for f, caso in sorted(_resumen.items())
                                for e in caso["ejemplos"]
                            ]),
                            use_container_width=True, hide_index=True,
                        )

            # Diagnóstico agregado del maestro (solo conteos)
            diag = st.session_state.get("diagnostico_maestro") or {}
            if diag:
                with st.expander("🧪 Diagnóstico agregado del maestro", expanded=True):
                    d1, d2, d3 = st.columns(3)
                    d1.metric("Filas procesadas", diag["filas"])
                    d2.metric("Columnas del Excel", diag["columnas"])
                    d3.metric("Filas descartadas", diag["filas_eliminadas"])

                    if diag["ausentes"]:
                        st.warning(
                            "Columnas que NO vienen en el archivo (su regla no se aplicó): "
                            + " · ".join(diag["ausentes"])
                        )
                    else:
                        st.success("Todas las columnas de las 14 reglas vienen en el archivo.")

                    st.markdown("**Ocupación por columna (antes de aplicar las reglas)**")
                    ocup = pd.DataFrame(diag["ocupacion"])
                    if not ocup.empty:
                        detalle_m = st.session_state.get("detalle_reglas") or {}
                        cambios_por_col = {
                            m_col: m["cambios_reales"]
                            for k, m in detalle_m.items()
                            for m_col in [k.split(". ", 1)[-1].rsplit(" (", 1)[0]]
                        }
                        ocup["Cambios reales"] = ocup["columna"].map(cambios_por_col).fillna(0).astype(int)
                        ocup = ocup.rename(columns={
                            "regla": "Regla", "columna": "Columna",
                            "con_valor": "Con valor", "vacias": "Vacías",
                        })
                        st.dataframe(ocup, use_container_width=True, hide_index=True)

                    # Texto plano copiable, sin datos personales
                    lineas = [
                        f"Filas procesadas: {diag['filas']} (originales {diag['filas_originales']}, "
                        f"descartadas {diag['filas_eliminadas']})",
                        f"Columnas del Excel: {diag['columnas']}",
                        f"Columnas ausentes: {', '.join(diag['ausentes']) or 'ninguna'}",
                        "",
                        f"{'COLUMNA':52s} {'CON VALOR':>10s} {'VACIAS':>8s} {'CAMBIOS':>8s}",
                    ]
                    for _fila in diag["ocupacion"]:
                        _camb = int(ocup.loc[ocup["Columna"] == _fila["columna"], "Cambios reales"].iloc[0]) \
                                if not ocup.empty else 0
                        lineas.append(
                            f"{_fila['columna'][:52]:52s} {_fila['con_valor']:10d} "
                            f"{_fila['vacias']:8d} {_camb:8d}"
                        )
                    total_cambios = sum(m["cambios_reales"] for m in (st.session_state.get("detalle_reglas") or {}).values())
                    lineas += ["", f"TOTAL cambios reales: {total_cambios}"]
                    st.caption("Resumen copiable (solo conteos, sin datos personales):")
                    st.code("\n".join(lineas), language="text")

            # Verificación post-proceso de las reglas 7 y 8 (Fonasa)
            verif = st.session_state.get("verificacion_fonasa") or {}
            if verif.get("filas_fonasa"):
                with st.expander(
                    f"✅ Verificación reglas 7-8 (Fonasa) — {verif['filas_fonasa']} fila(s) Fonasa",
                    expanded=True,
                ):
                    st.caption(
                        "Valores que quedaron en las columnas de monto DESPUÉS de aplicar "
                        "el override de Fonasa. La columna UF debe mostrar solo '%'."
                    )
                    for _col, _conteo in verif.items():
                        if _col == "filas_fonasa":
                            continue
                        st.markdown(f"**{_col}**")
                        st.dataframe(
                            pd.DataFrame(
                                [{"Valor final": k, "Filas Fonasa": v} for k, v in _conteo.items()]
                            ),
                            use_container_width=True, hide_index=True,
                        )

            # Detalle por regla nueva (separado de los defaults preexistentes)
            detalle = st.session_state.get("detalle_reglas") or {}
            if detalle:
                with st.expander(
                    f"⚙️ Detalle de las reglas nuevas "
                    f"({sum(m['cambios_reales'] for m in detalle.values())} celda(s) modificada(s))",
                    expanded=False,
                ):
                    st.caption(
                        "**Cambios reales** = celdas cuyo valor es distinto antes vs. después "
                        "(no filas procesadas). *Rellenos* = la celda estaba vacía. "
                        "*Sobrescrituras* = la celda tenía un valor y se reemplazó. "
                        "*No vacías antes* dice si la columna venía con datos del cliente."
                    )
                    st.dataframe(
                        pd.DataFrame([
                            {
                                "Regla":            k,
                                "Filas":            m["filas"],
                                "No vacías antes":  m["no_vacias_antes"],
                                "Cambios reales":   m["cambios_reales"],
                                "Rellenos":         m["rellenos"],
                                "Sobrescrituras":   m["sobrescrituras"],
                            }
                            for k, m in sorted(detalle.items())
                        ]),
                        use_container_width=True, hide_index=True,
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
            reporte_txt += f"Filas eliminadas (no A/F): {filas_eliminadas}\n"
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
