# -*- coding: utf-8 -*-
"""
Migración de empleados Rex -> BNOVUS
====================================
Página de rex-tools que toma el "Listado de Empleados" exportado desde Rex y
genera el archivo de carga de trabajadores en el formato de la plantilla BNOVUS.

Insumos que sube el usuario:
  - Listado de Empleados de Rex (.xlsx). Header en la fila 2, datos desde la fila 3.

Parámetros ingresables (barra lateral):
  - RUT Empresa            (obligatorio; columna A de BNOVUS)
  - Alcance                (todos / solo activos)
  - Defaults de negocio    (moneda sueldo, tipo sueldo, cantidad días, modalidad,
                            gratificación) para las columnas que Rex NO trae.

Salida:
  - Archivo .xlsx con la estructura BNOVUS (hoja Sheet1 + catálogos), listo para subir.
  - Informe de cobertura: valores no reconocidos en los mapeos (para revisar).

La plantilla BNOVUS vive en  data/plantilla_bnovus.xlsx  (se puede reemplazar por
file_uploader si no está en el repo).
"""

import io
import os
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st
import openpyxl

# ----------------------------------------------------------------------------- #
#  Config de página + branding rex-tools (igual que las demás páginas)
# ----------------------------------------------------------------------------- #
st.set_page_config(page_title="Migración BNOVUS | Rex+ Tools",
                   page_icon="👥", layout="wide")
try:
    from lib.branding import aplicar_branding, aplicar_footer, hero
    BRANDING = True
except ImportError:
    BRANDING = False

if BRANDING:
    aplicar_branding(titulo_pagina="Migración BNOVUS", badge="BETA")

TEMPLATE_PATH = os.path.join("data", "plantilla_bnovus.xlsx")

# ----------------------------------------------------------------------------- #
#  Utilidades de normalización
# ----------------------------------------------------------------------------- #
def _norm(s):
    """minúsculas, sin acentos, sin espacios extremos -> para comparar llaves."""
    if s is None:
        return ""
    s = str(s).strip()
    s = "".join(c for c in unicodedata.normalize("NFD", s)
                if unicodedata.category(c) != "Mn")
    return s.lower()


def _clean(s):
    if s is None:
        return ""
    return str(s).strip()


def _to_date(v):
    """Devuelve datetime o None."""
    if v is None or v == "" or (isinstance(v, str) and v.strip() == ""):
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, (int, float)):
        # serial excel improbable aquí; ignorar
        return None
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(str(v).strip()[:10], fmt)
        except ValueError:
            continue
    return None


# ----------------------------------------------------------------------------- #
#  Catálogos de mapeo  (Rex -> BNOVUS)
#  Las llaves se comparan normalizadas (_norm), así toleran mayúsculas/acentos.
# ----------------------------------------------------------------------------- #
SEXO_MAP = {"m": "MASCULINO", "f": "FEMENINO"}

# Formato según archivo BNOVUS aceptado: sin "(a)", en título.
ESTCIVIL_MAP = {
    "s": "Soltero",
    "c": "Casado",
    "d": "Divorciado",
    "v": "Viudo",
    "u": "Conviviente Civil",   # 'U' = unión/conviviente civil (confirmar con negocio)
}

REGION_MAP = {
    _norm("Antofagasta"): "Región de Antofagasta",
    _norm("Metropolitana de Santiago"): "Región Metropolitana",
    _norm("Libertador General Bernardo O'Higgins"):
        "Región del Libertador General Bernardo O Higgins",
    _norm("Valparaíso"): "Región de Valparaiso",
    _norm("Coquimbo"): "Región de Coquimbo",
    _norm("Tarapacá"): "Región de Tarapacá",
    _norm("Arica y Parinacota"): "Región de Arica y Parinacota",
    _norm("Biobío"): "Región del Bío-Bío",
    _norm("Maule"): "Región del Maule",
    _norm("Ñuble"): "Región de Ñuble",
    _norm("De los Lagos"): "Región de los Lagos",
    _norm("Los Lagos"): "Región de los Lagos",
    _norm("Atacama"): "Región de Atacama",
    _norm("La Araucanía"): "Región del Araucanía",
    _norm("Araucanía"): "Región del Araucanía",
    _norm("Los Ríos"): "Región de los Ríos",
    _norm("Magallanes"): "Región de Magallanes y la Antartica Chilena",
    _norm("Aysén"): "Región de Aysén del General Carlos Ibáñez del Campo",
}

FORMAPAGO_MAP = {
    "actacorr": "Abono en CuentaCte",
    "actavis": "Abono en Cuenta Vista",
    "actaaho": "Abono en CuentaAhorro",
    "actarut": "Cuenta RUT",
    "efectivo": "Efectivo",
    "cheque": "Cheque",
    "sindefinir": "",
}

# Formato según archivo BNOVUS aceptado: en MAYÚSCULAS.
TIPOCONTRATO_MAP = {
    "i": "INDEFINIDO",
    "f": "A PLAZO FIJO",
    "o": "POR OBRA O FAENA",
    "h": "HONORARIOS",
}

# Jubilado? -> Prevision Tipo Trabajador (catálogo TipoTrabAFP)
PREVTIPO_MAP = {
    _norm("Activo (No Pensionado)"): "Activo",
    _norm("Pensionado y no cotiza"): "Pensionado (no cotiza)",
    _norm("Pensionado y cotiza"): "Pensionado (cotiza)",
    _norm("Activo > 60 ó 65 años"): "Activo > 60 ó 65 años",
}

ESTADO_MAP = {"a": "V", "p": "D"}   # Estado contrato Rex -> Estado Empleado BNOVUS

# AFP: se pasan a MAYÚSCULAS; casos con nombre distinto van acá
AFP_MAP = {
    "afp": "",   # genérico sin definir
}

ISAPRE_MAP = {
    "fonasa": "FONASA",
    "nuevamasvida": "NUEVA MASVIDA",
    "banmedica": "BANMEDICA",
    "colmena": "COLMENA",
    "consalud": "CONSALUD",
    "cruzblanca": "CRUZ BLANCA",
    "esencial": "ESENCIAL",
    "vidatres": "VIDA TRES",
    "fundacion": "FUNDACION",
}

# Banco (código Rex -> nombre banco BNOVUS). Best-effort: ajustar a catálogo BNOVUS.
BANCO_MAP = {
    "estado": "BANCO ESTADO",
    "falabella": "BANCO FALABELLA",
    "chile": "BANCO DE CHILE",
    "santander": "BANCO SANTANDER",
    "bci": "CREDITO E INVERSIONES",
    "scotia": "SCOTIABANK",
    "itau": "ITAU",
    "mercadopago": "MERCADO PAGO",
    "ripley": "BANCO RIPLEY",
    "losandes": "CAJA LOS ANDES",
    "security": "BANCO SECURITY",
    "tenpo": "TENPO",
    "copeuch": "COOPEUCH",
    "bice": "BANCO BICE",
    "edwards": "BANCO EDWARDS",
    "bbva": "BBVA",
    "corpbanca": "CORPBANCA",
}

# País (Rex) -> Nacionalidad (título), según archivo BNOVUS aceptado.
NACIONALIDAD_MAP = {
    "chile": "Chilena", "chilena": "Chilena",
    "peru": "Peruana", "peruana": "Peruana",
    "bolivia": "Boliviana", "boliviana": "Boliviana",
    "venezuela": "Venezolana", "venezolana": "Venezolana",
    "colombia": "Colombiana", "colombiana": "Colombiana",
    "argentina": "Argentina",
    "ecuador": "Ecuatoriana", "ecuatoriana": "Ecuatoriana",
    "haiti": "Haitiana", "haitiana": "Haitiana",
    "brasil": "Brasileña",
}

# Fecha término placeholder para contratos indefinidos (archivo aceptado usa 31-12-2030).
FEC_TERMINO_INDEFINIDO = datetime(2030, 12, 31)

# Índice de columnas BNOVUS por encabezado exacto de la fila 1 de la plantilla
BNOVUS_HEADERS = [
    "Rut Empresa", "Codigo Interno Trabajador", "Rut Trabajador", "Nombre Trabajador",
    "Apellido Paterno Trabajador", "Apellido Materno Trabajador", "fecha nac",
    "Genero Trabajador", "Nacionalidad Trabajador", "Estado Civil Trabajador",
    "Email Personal Trabajador", "Rut Jefe Directo Trabajador", "Sindicato Trabajador",
    "Area nivel 1", "Area nivel 2", "Area nivel 3", "Area nivel 4", "Sucursal Trabajador",
    "Cargo Trabajador", "Email Corporativo Trabajador", "Direccion Particular Trabajador",
    "Direccion Nro. Casa Trabajador", "Direccion Nro. Depto. Trabajador", "Comuna Trabajador",
    "Region Trabajador", "Tipo Contrato", "Fecha Ingreso Trabajador", "Fecha Firma Contrato",
    "Modalidad Contrato", "Sueldo base Contrato", "Centro de Costo Trabajador",
    "Codigo Centro Costo", "Fecha Inicio Contrato", "Fecha Termino Contrato",
    "Fecha Devengacion Vacaciones", "Tipo Sueldo Contrato", "Moneda Contrato",
    "Horas por Semana Contrato", "Cantidad Dias Contrato", "Tipo Gratificacion",
    "Estado Empleado", "Valor Gratificacion", "Monto Movilizacion", "Monto Colacion",
    "Monto Anticipo", "Tipo Forma de Pago", "Banco Trabajador",
    "Numero Cta. Corriente Trabajador", "Prevision Trabajador", "Prevision Tipo Trabajador",
    "Institucion de Salud", "Modalidad de Pactado", "Cotizacion en Pesos de Pactado",
    "Cotizacion en UF de Pactado", "Habilitar Cotizacion Voluntaria",
    "Moneda Cotizacion Voluntaria", "Monto Cotizacion Voluntaria",
    "Rebaja Trib Art42 Cotizacion Voluntaria", "Instit Admin Cotizacion Voluntaria",
    "Habilitar Seguro Cesantia", "Fecha Ingreso Seguro Cesantia",
    "Fecha Termino Seguro Cesantia", "Fecha Ultima Cotizacion Seguro Cesantia",
    "Afp Seguro Cesantia", "Tipo Seguro de Vida", "Aseguradora Seguro de Vida",
    "Poliza Seguro de Vida", "Beneficiarios Seguro de Vida", "Oficinadireccion laboral",
    "Oficina laboral", "Oficinapiso laboral", "Oficinaanexo laboral", "Oficinacomuna laboral",
    "Username Trabajador", "Areacodigo Trabajador", "Codigocargo Trabajador",
    "Numero Contrato", "Codigo JornadaEspecial", "Jornada Especial", "Rol", "Grupo",
    "Lista1", "Lista2", "Lista3", "Lista4", "Lista5", "Lista6", "Lista7", "Lista8",
    "Lista9", "Lista10", "TextoAdic1", "TextoAdic2", "FechaAdicional1", "FechaAdicional2",
    "Grado Educacional", "Carrera1", "Institucion Educacion Superior 1",
    "Ultimo cargo trabajado", "Ultima Empresa",
]
COL = {h: i for i, h in enumerate(BNOVUS_HEADERS)}   # nombre -> índice 0-based


# ----------------------------------------------------------------------------- #
#  Split de nombre "APELLIDO_P APELLIDO_M NOMBRES..."
# ----------------------------------------------------------------------------- #
def split_nombre(nombre):
    toks = _clean(nombre).split()
    if not toks:
        return "", "", ""
    if len(toks) == 1:
        return "", toks[0], ""           # solo un token -> paterno
    if len(toks) == 2:
        return "", toks[0], toks[1]      # paterno, materno (sin nombres)
    paterno, materno = toks[0], toks[1]
    nombres = " ".join(toks[2:])
    return nombres, paterno, materno


# ----------------------------------------------------------------------------- #
#  Transformación principal
# ----------------------------------------------------------------------------- #
def transformar(df, rut_empresa, incluir_todos, defaults, avisos):
    """
    df       : DataFrame del Listado de Empleados de Rex (columnas = header fila 2)
    devuelve : (lista_de_filas, dict_no_mapeados)
    """
    no_map = {"region": set(), "afp": set(), "isapre": set(), "banco": set(),
              "estado_civil": set(), "forma_pago": set(), "tipo_contrato": set(),
              "prev_tipo": set()}

    def g(row, *names):
        """primer valor no nulo de las columnas 'names' (tolerante a duplicados .1)."""
        for n in names:
            if n in row.index and pd.notna(row[n]):
                return row[n]
        return None

    filas = []
    for _, row in df.iterrows():
        rut = _clean(g(row, "Rut"))
        if not rut:
            continue

        estado_src = _norm(g(row, "Estado"))
        estado_bn = ESTADO_MAP.get(estado_src, "V")
        if not incluir_todos and estado_bn == "D":
            continue

        nombres, paterno, materno = split_nombre(g(row, "Nombre"))

        # --- códigos con catálogo ---
        sexo = SEXO_MAP.get(_norm(g(row, "Sexo")), _clean(g(row, "Sexo")).upper())

        ec_src = _norm(g(row, "Estado civil"))
        ec = ESTCIVIL_MAP.get(ec_src)
        if ec is None and ec_src:
            no_map["estado_civil"].add(_clean(g(row, "Estado civil")))
            ec = ""

        reg_src = _norm(g(row, "Región"))
        reg = REGION_MAP.get(reg_src)
        if reg is None and reg_src:
            no_map["region"].add(_clean(g(row, "Región")))
            reg = _clean(g(row, "Región"))

        fp_src = _norm(g(row, "Forma Pago"))
        fp = FORMAPAGO_MAP.get(fp_src)
        if fp is None and fp_src:
            no_map["forma_pago"].add(_clean(g(row, "Forma Pago")))
            fp = ""

        afp_src = _norm(g(row, "AFP"))
        afp = AFP_MAP.get(afp_src, _clean(g(row, "AFP")).upper())

        isa_src = _norm(g(row, "Isapre"))
        isa = ISAPRE_MAP.get(isa_src)
        if isa is None and isa_src:
            no_map["isapre"].add(_clean(g(row, "Isapre")))
            isa = _clean(g(row, "Isapre")).upper()

        banco_src = _norm(g(row, "Banco"))
        banco = BANCO_MAP.get(banco_src)
        if banco is None and banco_src:
            no_map["banco"].add(_clean(g(row, "Banco")))
            banco = _clean(g(row, "Banco")).upper()

        tc_src = _norm(g(row, "Tipo contr."))
        tc = TIPOCONTRATO_MAP.get(tc_src)
        if tc is None and tc_src:
            no_map["tipo_contrato"].add(_clean(g(row, "Tipo contr.")))
            tc = ""

        prev_src = _norm(g(row, "Jubilado?"))
        prev_tipo = PREVTIPO_MAP.get(prev_src, "Activo")

        nacion = NACIONALIDAD_MAP.get(_norm(g(row, "País")),
                                      _clean(g(row, "País")).title() or "Chilena")

        # --- salud: modalidad de pactado ---
        moneda_isa = _norm(g(row, "Moneda Isapre"))
        cot_uf = g(row, "Cotización UF")
        cot_pesos = g(row, "Cotización $")
        if moneda_isa in ("u.f.", "uf"):
            modalidad_pactado = "UF"
            salud_uf = cot_uf if (cot_uf not in (None, 0)) else ""
            salud_pesos = ""
        else:   # 7% (Fonasa u opción legal)
            modalidad_pactado = 0.07           # opción "7%" del catálogo TipoPactoSalud
            salud_uf = ""
            salud_pesos = ""

        # --- fechas ---
        fec_ini_contr = _to_date(g(row, "Fecha Inicio contrato"))
        fec_term_contr = _to_date(g(row, "Fecha término contrato"))
        # placeholder de "indefinido" (año >= 2999 en Rex) -> ignorar
        if fec_term_contr is not None and fec_term_contr.year >= 2999:
            fec_term_contr = None
        if tc == "INDEFINIDO":
            # archivo BNOVUS aceptado usa 31-12-2030 como término de indefinidos
            fec_term_bn = FEC_TERMINO_INDEFINIDO
        else:
            fec_term_bn = fec_term_contr
        fec_venc_vac = _to_date(g(row, "Fecha inicio vacaciones")) or fec_ini_contr
        fec_cesantia = _to_date(g(row, "Fecha inc. Seguro Cesa.")) or fec_ini_contr

        # --- seguro cesantía ---
        afecto_ces = g(row, "Afecto Seguro Cesantéa", "Afecto Seguro Cesantia")
        habilita_ces = "S" if bool(afecto_ces) else "N"

        # --- código interno: la columna "Código Interno" de Rex trae basura
        #     (valores como 'SI', 'CONTRATO 1', códigos de centro de costo),
        #     y el archivo BNOVUS aceptado la deja vacía -> siempre en blanco.
        cod_int = ""

        # --- sindicato ---
        sind_src = _norm(g(row, "Sindicato"))
        if sind_src in ("", "no tiene", "sindivpten", "no sindicalizado",
                        "no sindicalizados"):
            sindicato = "No Sindicalizados"
        else:
            sindicato = _clean(g(row, "Sindicato"))

        # ------------------------------------------------------------------ #
        #  Armado de la fila BNOVUS
        # ------------------------------------------------------------------ #
        fila = [None] * len(BNOVUS_HEADERS)

        def s(header, value):
            fila[COL[header]] = value

        s("Rut Empresa", rut_empresa)
        s("Codigo Interno Trabajador", cod_int)
        s("Rut Trabajador", rut)
        s("Nombre Trabajador", nombres)
        s("Apellido Paterno Trabajador", paterno)
        s("Apellido Materno Trabajador", materno)
        s("fecha nac", _to_date(g(row, "Fecha Nacimiento")))
        s("Genero Trabajador", sexo)
        s("Nacionalidad Trabajador", nacion)
        s("Estado Civil Trabajador", ec)
        s("Email Personal Trabajador",
          _clean(g(row, "Email Personal", "Correo electrónico")))
        s("Sindicato Trabajador", sindicato)
        # organigrama no viene en Rex -> GENERAL en nivel 1 (regla plantilla)
        s("Area nivel 1", defaults["area_nivel1"])
        area_src = _clean(g(row, "Área"))
        s("Area nivel 2", area_src)
        s("Sucursal Trabajador", _clean(g(row, "Sede")))
        s("Cargo Trabajador", _clean(g(row, "Cargo")))
        s("Email Corporativo Trabajador", "")
        s("Direccion Particular Trabajador", _clean(g(row, "Dirección")))
        s("Comuna Trabajador", _clean(g(row, "Comuna")))
        s("Region Trabajador", reg)
        s("Tipo Contrato", tc)
        s("Fecha Ingreso Trabajador", fec_ini_contr)
        s("Fecha Firma Contrato", fec_ini_contr)
        s("Modalidad Contrato", defaults["modalidad_contrato"])
        s("Sueldo base Contrato", g(row, "Sueldo Base"))
        s("Centro de Costo Trabajador", _clean(g(row, "Centro Costo")))
        s("Codigo Centro Costo", _clean(g(row, "Id Centro de Costo")))
        s("Fecha Inicio Contrato", fec_ini_contr)
        s("Fecha Termino Contrato", fec_term_bn)
        s("Fecha Devengacion Vacaciones", fec_venc_vac)
        s("Tipo Sueldo Contrato", defaults["tipo_sueldo"])
        s("Moneda Contrato", defaults["moneda_sueldo"])
        s("Horas por Semana Contrato", g(row, "Horas Semanales"))
        s("Cantidad Dias Contrato", defaults["cantidad_dias"])
        s("Tipo Gratificacion", defaults["tipo_gratificacion"])
        s("Estado Empleado", estado_bn)
        s("Valor Gratificacion", defaults["valor_gratificacion"])
        s("Monto Movilizacion", g(row, "Movilización"))
        s("Monto Colacion", g(row, "Colación"))
        s("Monto Anticipo", "SIN ANTICIPO")
        s("Tipo Forma de Pago", fp)
        s("Banco Trabajador", banco)
        s("Numero Cta. Corriente Trabajador", _clean(g(row, "Cuenta Banco")))
        s("Prevision Trabajador", afp)
        s("Prevision Tipo Trabajador", prev_tipo)
        s("Institucion de Salud", isa)
        s("Modalidad de Pactado", modalidad_pactado)
        s("Cotizacion en Pesos de Pactado", salud_pesos)
        s("Cotizacion en UF de Pactado", salud_uf)
        s("Habilitar Seguro Cesantia", habilita_ces)
        s("Fecha Ingreso Seguro Cesantia", fec_cesantia)
        s("Afp Seguro Cesantia", afp)
        s("Numero Contrato", 1)
        s("Rol", "empleados")
        s("Grupo", "Todos los empleados")

        filas.append(fila)

    return filas, no_map


DATE_HEADERS = {
    "fecha nac", "Fecha Ingreso Trabajador", "Fecha Firma Contrato",
    "Fecha Inicio Contrato", "Fecha Termino Contrato", "Fecha Devengacion Vacaciones",
    "Fecha Ingreso Seguro Cesantia", "Fecha Termino Seguro Cesantia",
    "Fecha Ultima Cotizacion Seguro Cesantia", "FechaAdicional1", "FechaAdicional2",
}


def construir_workbook(filas, template_bytes):
    """Escribe las filas en la plantilla BNOVUS y devuelve bytes del .xlsx."""
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws = wb["Sheet1"]
    # limpiar datos previos (por si la plantilla trae ejemplo)
    for r in range(2, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    date_cols = {COL[h] for h in DATE_HEADERS if h in COL}
    for i, fila in enumerate(filas, start=2):
        for j, val in enumerate(fila):
            cell = ws.cell(i, j + 1)
            cell.value = val
            if j in date_cols and isinstance(val, datetime):
                cell.number_format = "DD-MM-YYYY"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


# ----------------------------------------------------------------------------- #
#  UI
# ----------------------------------------------------------------------------- #
def main():
    if BRANDING:
        hero("👥 Migración de empleados · Rex → BNOVUS",
             "Convierte el Listado de Empleados de Rex al archivo de carga de "
             "trabajadores de BNOVUS.")
    else:
        st.title("Migración de empleados · Rex → BNOVUS")
        st.caption("Convierte el Listado de Empleados de Rex al archivo de carga de "
                   "trabajadores de BNOVUS.")

    # --- Parámetros en el cuerpo principal (siempre visibles) ---
    c1, c2 = st.columns([2, 1])
    with c1:
        rut_empresa = st.text_input(
            "RUT Empresa *",
            help="Sin puntos, con guión y dígito verificador. Ej: 76361420-4",
            placeholder="76361420-4",
        ).strip()
    with c2:
        alcance = st.radio(
            "Alcance", ["Todos", "Solo activos"], horizontal=True,
            help="Solo activos excluye a los trabajadores con estado 'P'.",
        )

    with st.expander("Parámetros avanzados (valores por defecto)"):
        d1, d2, d3 = st.columns(3)
        with d1:
            moneda_sueldo = st.selectbox("Moneda Contrato",
                                         ["Peso", "UF", "Dolar", "Euro", "UTM"], index=0)
            tipo_sueldo = st.selectbox("Tipo Sueldo Contrato",
                                       ["Sueldo Privado", "Sueldo Público"], index=0)
        with d2:
            modalidad_contrato = st.selectbox(
                "Modalidad Contrato",
                ["Con Horario", "Sin Horario", "Honorarios"], index=0)
            cantidad_dias = st.number_input("Cantidad Dias Contrato", 1, 7, 5)
        with d3:
            tipo_gratificacion = st.selectbox("Tipo Gratificación",
                                              ["", "Calculada", "Fija"], index=1)
            valor_gratificacion = st.text_input("Valor Gratificación", value="TOPE 4,75")
        area_nivel1 = st.text_input("Area nivel 1 (organigrama)", value="GENERAL")

    archivo = st.file_uploader("Listado de Empleados de Rex (.xlsx)", type=["xlsx"])

    tpl_bytes = None
    if os.path.exists(TEMPLATE_PATH):
        with open(TEMPLATE_PATH, "rb") as f:
            tpl_bytes = f.read()
    else:
        st.info("No se encontró data/plantilla_bnovus.xlsx. Súbela manualmente:")
        tpl_up = st.file_uploader("Plantilla BNOVUS (.xlsx)", type=["xlsx"],
                                  key="tpl")
        if tpl_up is not None:
            tpl_bytes = tpl_up.read()

    if st.button("Generar archivo BNOVUS", type="primary", disabled=archivo is None):
        if not rut_empresa:
            st.error("Ingresa el RUT Empresa antes de generar.")
            st.stop()
        if tpl_bytes is None:
            st.error("Falta la plantilla BNOVUS.")
            st.stop()

        # header en la fila 2 (skiprows=1). Título en la fila 1.
        df = pd.read_excel(archivo, sheet_name=0, skiprows=1)
        df = df[df["Rut"].notna()]

        defaults = dict(
            moneda_sueldo=moneda_sueldo, tipo_sueldo=tipo_sueldo,
            modalidad_contrato=modalidad_contrato, cantidad_dias=cantidad_dias,
            tipo_gratificacion=tipo_gratificacion,
            valor_gratificacion=valor_gratificacion, area_nivel1=area_nivel1,
        )
        avisos = []
        filas, no_map = transformar(df, rut_empresa,
                                    alcance == "Todos", defaults, avisos)

        data = construir_workbook(filas, tpl_bytes)

        st.success(f"Archivo generado: {len(filas)} trabajadores.")
        rut_slug = rut_empresa.replace("-", "").replace(".", "")
        st.download_button(
            "⬇️ Descargar archivo BNOVUS",
            data=data,
            file_name=f"bnovus_{rut_slug}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # informe de cobertura
        pend = {k: v for k, v in no_map.items() if v}
        if pend:
            st.warning("Valores no reconocidos en los catálogos (se copiaron tal cual, "
                       "revisar):")
            for k, v in pend.items():
                st.write(f"**{k}**: {', '.join(sorted(map(str, v)))}")
        else:
            st.info("Todos los valores codificados se mapearon correctamente.")

    if BRANDING:
        aplicar_footer()


if __name__ == "__main__" or True:
    main()
