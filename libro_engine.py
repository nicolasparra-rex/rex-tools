# -*- coding: utf-8 -*-
"""Motor de lectura de 'libro de remuneraciones' de cualquier cliente -> mapeo a conceptos Rex.
Autodetección de estructura + clasificación por bloque (posición) + propuesta de ID por
sinónimos y catálogo del cliente. Suma columnas que van al mismo concepto. Cuadra al peso."""
import unicodedata, calendar, pandas as pd
from collections import OrderedDict

def norm(s):
    if s is None: return ""
    return "".join(c for c in unicodedata.normalize("NFD", str(s).strip().lower()) if unicodedata.category(c) != "Mn")

def _num(v):
    return v if isinstance(v, (int, float)) and pd.notna(v) else 0

# ---------- Columnas ESTRUCTURALES (no son conceptos) ----------
STRUCT = {
 "rut": ["numero de documento","rut trabajador","rut del trabajador","rut","n de documento","n documento"],
 "nombre": ["nombre completo","nombre","apellido y nombre","nombre trabajador"],
 "dias_trab": ["dias trabajados"],
 "afp": ["fondo de cotizacion","afp","prevision"],
 "salud": ["fonasa/isapre","salud","isapre"],
 "base_afp": ["base imponible afp","imponible afp","imp. prev./salud","imponible","imponible topeado"],
 "base_ces": ["base imponible cesantia","imponible cesantia","imp. cesantia","imponible seguro cesantia"],
 "base_trib": ["base tributable","tributable","afecto impuesto"],
 "total_haberes": ["total haberes"],
 "total_descuentos": ["total descuentos"],
 "total_aportes": ["total aportes"],
 "liquido": ["liquido a pago","liquido a recibir","liquido"],
 "plan_uf": ["plan isapre uf","plan isapre"],
 "tramo": ["tramo"],
 "sueldo_pactado": ["sueldo contractual","sueldo base pactado","sueldo pactado"],
 "fecha_ingreso": ["fecha de ingreso","fecha ingreso compania","fecha ingreso","fecha de alta","inicio contrato"],
}
IGNORAR = {"sueldo","plan uf","fecha de baja","departamento","tipo de empleado","id centro de costo",
 "nombre c. costo","area","sede","workday - grade","grade","periodo","empleado","workday - nombre",
 "workday - role","workday - id","workday - centro de costo - codigo","workday - centro de costo - descripcion",
 "workday - centro de costo -\ndescripcion","nombre","centro de costo","cargo","categoria",
 "workday - centro de costo"}

# ---------- Diccionario de sinónimos de CONCEPTOS (nombre normalizado -> id Rex) ----------
CONCEPTO = {
 "sueldo base":"sueldoBase","sueldo":"sueldoBase",
 "gratificacion":"gratificacion","diferencia de gratificacion":"gratificacion","diferencia gratificacion":"gratificacion",
 "bono ayuda hijo menor":"BonoHijoMenor","bono ayuda hijo menor (l)":"BonoHijoMenor",
 "bono internet celular":"BonoCelular","bono internet - celular":"BonoCelular","bonocelular":"BonoCelular",
 "bono sala cuna":"BonoSalaCunaG","dif subsidio licencia medica":"DifSubLicenciaMedica",
 "dif. subsidio lic. medica":"DifSubLicenciaMedica",
 "asignacion de alimentacion":"AsigAlimentacion","asignacion alimentacion":"AsigAlimentacion",
 "asignacion de movilizacion":"movilizacion","asignacion movilizacion":"movilizacion","movilizacion":"movilizacion",
 "asignacion de teletrabajo":"AsignacionTeletrabjo","asignacion teletrabajo":"AsignacionTeletrabjo","diferencia a pagar":"DiferenciaAPagar",
 "guardia pasiva":"GuardiaPasiva","hora extra guardia activa":"HEGuardiaActiva",
 "horas extras al 50 %":"horasEx50","horas extras al 50%":"horasEx50",
 "indemnizacion sustitutiva previo aviso":"iasMes","mes de aviso":"iasMes","indemnizacion por vacaciones":"iasVacaciones",
 "indemnizacion feriado legal":"iasVacaciones","indemnizacion a. de servicio":"iasLegal",
 "indemnizacion legal anos de servicio":"iasLegal",
 "bono extraordinario":"BonoExtraordinario","bono extraordinario marketing":"BonoMarketing",
 "bono de demanda":"BonoDemanda","bono demanda":"Bono_Demanda","bono referidos (l)":"Bono_Referido",
 "aguinaldo":"AguinaldoMi","aguinaldo(l)":"AguinaldoMi","bono":"BonoMi","bono sicp q4":"BonoExtraordinario",
 "diferencia de sueldo":"DiferenciaAPagar","vacaciones progresivas":"vacacionesProg",
 "sobregiro":"__SOBREGIRO__",
 # descuentos
 "descuento afp":"afp","cotiz. previ. obligatoria":"afp","cotizacion obligatoria previsional":"afp","a.f.p.":"afp",
 "cotizacion salud":"isapre","cotiz. salud obligatoria":"isapre","cotizacion fonasa":"isapre","salud":"isapre",
 "cotizacion adicional isapre":"__ISAPRE_AD__","adicional salud":"__ISAPRE_AD__","adicional isapre":"__ISAPRE_AD__",
 "seguro de desempleo":"cesEmpleado","seguro cesantia":"cesEmpleado","seguro de cesantia":"cesEmpleado",
 "impuesto unico":"impuesto","impuesto":"impuesto",
 "cotiz. prev. voluntaria":"apvi","a.p.v.":"apvi","apv":"apvi","cuenta ahorro voluntario afp":"afpAhor",
 "apv fintual adm gral fondos":"apvi","apv afp capital (a)":"apvi","apv afp capital":"apvi",
 "apv consorcio seg. de vida (a)":"apvi","apv vida security (a)":"apvi",
 "cotizacion c.c.a.f.":"cajaComp","cotizacion ccaf":"cajaComp","retencion adicional 3%":"solidarioremu",
 "anticipo sueldo":"anticipo","anticipo de sueldo":"anticipo","anticipo finiquito":"AnticipoPrestamoMi",
 "credito personal caja los andes":"cajaCred","credito los andes":"cajaCred",
 "leasing (ahorro) caja los andes":"cajaLeas","seguro de vida caja los andes":"cajaVida","seguro los andes":"cajaVida",
 "cuenta 2 pesos":"afpAhor","cuota sindical":"CuotaSindical","descuento falp":"DescuentoFalp",
 "descuento techops":"DescuentoTechops","retencion judicial":"retencionJudicial",
 "diferencia a descontar mes anterior":"sobregiro_anterior",
 # aportes empleador
 "aporte mutual":"mutual","mutual empleador":"mutual","aporte patronal ley sanna":"mutual",
 "aporte sis":"sis","sis":"sis","cesantia empleador":"cesAporteCi","seguro cesantia empleador":"cesAporteCi",
 "aporte seg. desemp. cta. ind.":"cesAporteCi","aporte seg. des. fondo solid.":"cesAporteSol",
 "ap. emp. cap. individual":"aporteAFPemp","afp prevision empleador":"aporteAFPemp",
 "ap. emp. ss exp. de vida":"aporteFAPPCEV","cotizacion expectativa de vida":"aporteFAPPCEV",
 # reliquidaciones (consolidado)
 "cotiz. previ. por reliquidacion":"reliquidaAfp","cotiz. salud obligatoria por reliquidacion":"reliquidaIsapre",
 "seguro cesantia por reliquidacion":"reliquidaCesEmpl","impuesto por reliquidacion":"reliquidaImpuesto",
 "sis por reliquidacion":"reliquidaSis","cesantia (empleador) por reliquidacion":"reliquidaCesCi",
 "mutual por reliquidacion":"reliquidaMutual","afp prevision empleador por reliquidacion":"reliquidaAporteAFP",
 "cotizacion expectativa de vida por reliquidacion":"reliquidaAporteCEV",
}

RUT_TOKENS = ["numero de documento","rut trabajador","rut del trabajador","n de documento","rut","n documento"]
INST_AFP_CONC = {"apvi","apvc","apviConvenido","afpAhor"}
INST_CAJA_CONC = {"cajaCred","cajaLeas","cajaVida"}
RELIQ = {"reliquidaAfp","reliquidaIsapre","reliquidaCesEmpl","reliquidaImpuesto","reliquidaSis",
         "reliquidaMutual","reliquidaCesCi","reliquidaCesSol","reliquidaAporteAFP","reliquidaAporteCEV"}
RELIQ_APO = {"reliquidaSis","reliquidaMutual","reliquidaCesCi","reliquidaCesSol","reliquidaAporteAFP","reliquidaAporteCEV"}
APORTES = {"mutual","sis","cesAporteCi","cesAporteSol","aporteAFPemp","aporteFAPPCEV"}

# ---------- Homologación de instituciones (match exacto -> "like") ----------
def cargar_homologacion(path_or_file):
    """Lee la tabla de instituciones. Tolera dos formatos de encabezado:
    (ID Institución, Clasificación, Nombre Institución, Código Equivalente) ó
    (institucion, clasificacion, nombre, codigoPrev)."""
    df = pd.read_excel(path_or_file, header=None)
    hr = 0
    for r in range(min(6, len(df))):
        if any(norm(x) == "clasificacion" for x in df.iloc[r].values): hr = r; break
    hdr = [norm(x) for x in df.iloc[hr].values]
    def col(names, d):
        for nm in names:
            if nm in hdr: return hdr.index(nm)
        return d
    ci = col(["id institucion", "institucion"], 0)
    cc = col(["clasificacion"], 1)
    cn = col(["nombre institucion", "nombre"], 2)
    ce = col(["codigo equivalente", "codigoprev", "codigo"], 4)
    out = []
    for _, row in df.iloc[hr+1:].iterrows():
        idv = row[ci] if ci < len(row) else None
        if pd.isna(idv) or not str(idv).strip(): continue
        out.append({"id": str(idv).strip(), "clasif": norm(row[cc]) if cc < len(row) else "",
                    "nombre_n": norm(row[cn]) if cn < len(row) else "", "id_n": norm(idv),
                    "cod": norm(row[ce]) if ce < len(row) else ""})
    return out

def resolver_inst(valor, homolog, clasifs=None):
    """Resuelve el nombre/valor de institución del libro al ID Rex vía la tabla. None si no hay match."""
    if not valor or not homolog: return None
    n = norm(valor); nc = n.replace(" ", "").replace("_", "").replace(".", "").replace("-", "")
    cand = [h for h in homolog if (clasifs is None or h["clasif"] in clasifs)]
    for h in cand:  # exacto por id / nombre / codigo
        if n == h["id_n"] or nc == h["id_n"].replace(" ", "") or n == h["nombre_n"] or (h["cod"] and nc == h["cod"]):
            return h["id"]
    for h in cand:  # like por id
        idn = h["id_n"].replace(" ", "")
        if idn and (idn in nc or nc in idn): return h["id"]
    STOP = {"afp","isapre","caja","comp","compania","de","la","el","los","las","del","chile",
            "seguro","seguros","vida","sin","empresa","aporta","mutual","fondo","fondos","adm","gral",
            "instituto","universidad","catolica","profesional","salud","isl","ips"}
    for h in cand:  # like por token distintivo del nombre ("AFP Capital"->capital, "ISAPRE Banmedica"->banmedica)
        toks = [t for t in h["nombre_n"].split() if len(t) > 3 and t not in STOP]
        if any(t in n for t in toks): return h["id"]
    return None

import re as _re
MESES_ES = {"enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,"julio":7,
            "agosto":8,"septiembre":9,"setiembre":9,"octubre":10,"noviembre":11,"diciembre":12}
def detectar_periodo(df, filename=""):
    """Detecta AAAA-MM desde el nombre del archivo o el contenido de la hoja. '' si no puede."""
    def buscar(txt):
        t = norm(txt)
        m = _re.search(r"(20\d\d)[-_ ]?(0[1-9]|1[0-2])(?!\d)", t)
        if m: return f"{m.group(1)}-{m.group(2)}"
        for name, mm in MESES_ES.items():
            m = _re.search(name + r"\s+de\s+(20\d\d)", t) or _re.search(name + r"\s+(20\d\d)", t) or _re.search(r"(20\d\d)\s+" + name, t)
            if m: return f"{m.group(1)}-{mm:02d}"
        return ""
    p = buscar(filename or "")
    if p: return p
    for r in range(min(12, len(df))):
        for x in df.iloc[r].values:
            if x is None: continue
            p = buscar(str(x))
            if p: return p
    return ""

# ---------- Dotación: resolución RUT -> contrato / empresa / mutual ----------
def cargar_dotacion(path_or_file):
    """Lee la dotación (query): Id empleado(RUT), Numero de contrato, fechaInic, idempresa, Mutual, % mutual."""
    df = pd.read_excel(path_or_file, header=None)
    hr = 0
    for r in range(min(6, len(df))):
        row = [norm(x) for x in df.iloc[r].values]
        if "id empleado" in row or "numero de contrato" in row: hr = r; break
    hdr = [norm(x) for x in df.iloc[hr].values]
    def col(names, d):
        for nm in names:
            if nm in hdr: return hdr.index(nm)
        return d
    ci = col(["id empleado","rut","rut trabajador"], 0)
    cc = col(["numero de contrato","contrato"], 1)
    cf = col(["fechainic","fecha inicio","fecha inicio contrato","fecha de ingreso"], 2)
    cemp = col(["idempresa","empresa"], 6)
    cmut = col(["mutual"], 7)
    cpm = col(["% mutual","cotizacionmutu","cotizacion mutual"], 8)
    cca = col(["caja","ccaf","cod caja","codigo caja","idcaja","id caja"], -1)
    out = {}
    for _, row in df.iloc[hr+1:].iterrows():
        idv = row[ci] if ci < len(row) else None
        if pd.isna(idv) or not str(idv).strip(): continue
        rut = str(idv).replace(".", "").strip().upper()
        try: fi = pd.to_datetime(row[cf], dayfirst=True) if (cf < len(row) and not pd.isna(row[cf])) else None
        except Exception: fi = None
        out.setdefault(rut, []).append({
            "contrato": row[cc] if (cc < len(row) and not pd.isna(row[cc])) else 1,
            "fechaInic": fi,
            "empresa": str(row[cemp]).strip() if (cemp < len(row) and not pd.isna(row[cemp])) else "",
            "mutual": str(row[cmut]).strip() if (cmut < len(row) and not pd.isna(row[cmut])) else "",
            "pmutual": row[cpm] if (cpm < len(row) and not pd.isna(row[cpm])) else "",
            "caja": str(row[cca]).strip() if (0 <= cca < len(row) and not pd.isna(row[cca])) else "",
        })
    return out

def resolver_contrato(rut, fecha_ing, periodo, dot):
    """Resuelve el contrato del RUT (lógica del sitio): 1 contrato -> directo; varios -> por fecha
    de ingreso exacta, o el vigente al período; si no se puede, cae a 1 y se marca."""
    rows = dot.get(rut)
    if not rows:
        return {"contrato": 1, "empresa": "", "mutual": "", "pmutual": "", "caja": "", "ok": False, "motivo": "RUT no está en la dotación"}
    if len(rows) == 1:
        r = rows[0]
        return {"contrato": r["contrato"], "empresa": r["empresa"], "mutual": r["mutual"], "pmutual": r["pmutual"], "caja": r.get("caja",""), "ok": True, "motivo": ""}
    # varios contratos: match por fecha de ingreso exacta
    if fecha_ing is not None:
        try: fd = pd.to_datetime(fecha_ing).date()
        except Exception: fd = None
        if fd is not None:
            for r in rows:
                if r["fechaInic"] is not None and pd.to_datetime(r["fechaInic"]).date() == fd:
                    return {"contrato": r["contrato"], "empresa": r["empresa"], "mutual": r["mutual"], "pmutual": r["pmutual"], "caja": r.get("caja",""), "ok": True, "motivo": ""}
    # vigente al período: mayor fechaInic <= último día del mes
    try:
        y, m = map(int, str(periodo).split("-")[:2])
        fin = pd.Timestamp(y, m, calendar.monthrange(y, m)[1])
        cand = [r for r in rows if r["fechaInic"] is not None and pd.to_datetime(r["fechaInic"]) <= fin]
        if cand:
            r = max(cand, key=lambda r: pd.to_datetime(r["fechaInic"]))
            return {"contrato": r["contrato"], "empresa": r["empresa"], "mutual": r["mutual"], "pmutual": r["pmutual"], "caja": r.get("caja",""), "ok": True, "motivo": ""}
    except Exception:
        pass
    return {"contrato": 1, "empresa": "", "mutual": "", "pmutual": "", "caja": "", "ok": False, "motivo": "multi-contrato sin fecha para desambiguar (se usó 1)"}

OUT_COLS = ["Fecha de proceso","Id empleado","Número de contrato","Id del concepto",
"Monto del concepto","Afecto","Id de institución","Cotización de jubilación","Días de licencias",
"Días trabajados","Fecha de aplicación","Empresa","Total de rebajas por LLSS","Rentas no gravadas",
"Rebaja por zona extrema","Jornada","Días de vacaciones","Monto Init","Fase","parcial7","parcial8"]

def load_grid(path, sheet_hint="libro"):
    xls = pd.ExcelFile(path)
    sh = [s for s in xls.sheet_names if sheet_hint in s.lower()]
    sh = sh[0] if sh else xls.sheet_names[0]
    return pd.read_excel(path, sheet_name=sh, header=None), sh

def detect_header_row(df, **_):
    for r in range(min(15, len(df))):
        for v in [norm(x) for x in df.iloc[r].values]:
            if v in RUT_TOKENS or v.startswith("rut trabajador") or v == "rut":
                return r
    return 0

def match_struct(hdr):
    """Detecta columnas estructurales. Prioriza coincidencia EXACTA sobre 'empieza con'."""
    norms = {i: norm(h) for i, h in enumerate(hdr) if h}
    out, used = {}, set()
    for exact in (True, False):
        for campo, syns in STRUCT.items():
            if campo in out: continue
            for i, n in norms.items():
                if i in used: continue
                if (n in syns) if exact else any(n.startswith(s) for s in syns):
                    out[campo] = i; used.add(i); break
    return out

# ---------- Catálogo de conceptos del cliente (export de Rex "Lista de conceptos") ----------
TIPO_BLOQUE = {
    "haber afecto": "haber", "haber exento": "haber", "haber solo tributable": "haber",
    "haber afecto especial": "haber", "vacaciones": "haber",
    "descuento": "desc", "descuento legal": "desc",
    "aporte empleador": "aporte",
}
def tipo_a_bloque(tipo):
    """Traduce el 'Tipo' del catálogo a bloque haber/desc/aporte. None si es Dato/Valor Guardado
    (esos no son montos de nómina; caen al respaldo por posición)."""
    return TIPO_BLOQUE.get(norm(tipo))

def leer_catalogo_rex(path_or_file):
    """Lee el catálogo de conceptos exportado de Rex (hoja 'Lista de conceptos' u otra).
    Devuelve (by_id, name_to_id): by_id[id]={'nombre','tipo','bloque'} y name_to_id[norm(nombre)]=id."""
    xl = pd.ExcelFile(path_or_file)
    sheet = next((s for s in xl.sheet_names if "concepto" in norm(s)), xl.sheet_names[0])
    raw = pd.read_excel(path_or_file, sheet_name=sheet, header=None)
    hr = 0
    for r in range(min(8, len(raw))):
        vals = [norm(x) for x in raw.iloc[r].values]
        if "concepto" in vals and any("nombre" in v for v in vals): hr = r; break
    hdr = [norm(x) for x in raw.iloc[hr].values]
    ci = hdr.index("concepto") if "concepto" in hdr else 0
    ni = next((i for i, v in enumerate(hdr) if "nombre" in v), 2)
    ti = next((i for i, v in enumerate(hdr) if v == "tipo"), None)
    by_id, name_to_id = {}, {}
    for _, row in raw.iloc[hr+1:].iterrows():
        cid = row[ci]
        if pd.isna(cid) or not str(cid).strip(): continue
        cid = str(cid).strip()
        nom = "" if ni is None or pd.isna(row[ni]) else str(row[ni]).strip()
        tipo = "" if ti is None or pd.isna(row[ti]) else str(row[ti]).strip()
        by_id[cid] = {"nombre": nom, "tipo": tipo, "bloque": tipo_a_bloque(tipo)}
        if nom: name_to_id.setdefault(norm(nom), cid)
    return by_id, name_to_id

def cargar_base_estandar(path_or_file):
    """Conjunto de IDs de concepto ESTÁNDAR de Rex (base común), para etiquetar legal vs propio."""
    try:
        df = pd.read_excel(path_or_file)
        col = "Concepto" if "Concepto" in df.columns else df.columns[0]
        return {str(v).strip() for v in df[col].dropna() if str(v).strip()}
    except Exception:
        return set()

def classify_and_map(hdr, struct, catalog_names=None, saved=None, valid_ids=None):
    """Propone el ID Rex de cada columna del libro.
    Prioridad: catálogo del cliente (por nombre) > diccionario legal (solo si el ID existe en el catálogo).
    Si valid_ids se entrega, cualquier propuesta que NO esté en ese conjunto se descarta (queda pendiente),
    de modo que nunca se cuela un ID que no exista en el catálogo del cliente."""
    catalog_names = catalog_names or {}; saved = saved or {}
    th = struct.get("total_haberes"); td = struct.get("total_descuentos")
    struct_idx = set(struct.values())
    # Sinónimos estructurales para excluir por NOMBRE. Se excluyen afp/salud/imponibles porque un cliente
    # puede tener columnas de CONCEPTO llamadas así (ej. 'afp' = monto de cotización, no la institución);
    # esas columnas estructurales igual se saltan por su posición detectada (struct_idx).
    _NO_SYN = {"afp", "salud", "base_afp", "base_ces", "base_trib"}
    struct_syn = set(s for campo, syns in STRUCT.items() if campo not in _NO_SYN for s in syns)
    def _ok(cid): return (valid_ids is None) or (cid in valid_ids)
    # Algunos libros ya traen los IDs de Rex como encabezado (ej. 'sueldoBase', 'cesAporteSol').
    # Permitimos casar el header directamente contra un ID del catálogo (por nombre normalizado).
    id_norm = {norm(c): c for c in (valid_ids or [])}
    filas = []
    for i, h in enumerate(hdr):
        n = norm(h)
        if not h or i in struct_idx or n in IGNORAR or n in struct_syn: continue
        if th is not None and i < th: grupo = "haber"
        elif td is not None and th is not None and th < i < td: grupo = "descuento"
        elif td is not None and i > td: grupo = "aporte"
        else: grupo = "?"
        cid, fuente, conf = None, "SIN MAPEAR", "-"
        if n in saved and _ok(saved[n]):
            cid, fuente, conf = saved[n], "guardado", "alta"
        elif n in catalog_names:                                   # el catálogo del cliente manda
            cid, fuente, conf = catalog_names[n], "catálogo", "alta"
        elif n in id_norm:                                          # el header YA es un ID del catálogo
            cid, fuente, conf = id_norm[n], "id-catálogo", "alta"
        elif n in CONCEPTO:                                        # diccionario legal, solo si existe en el catálogo
            c = CONCEPTO[n]
            if c == "__SOBREGIRO__": c = "compensaSobre" if grupo == "haber" else "sobregiro_anterior"
            if c == "__ISAPRE_AD__": c = "isapre"
            if _ok(c): cid, fuente, conf = c, "diccionario", "media"
        filas.append({"col": i+1, "header": h, "grupo": grupo, "id_rex": cid, "fuente": fuente, "confianza": conf})
    return filas

def generar_detalle(df, header_row, struct, mapping, params_row, cot_hist, config, homolog=None, dotacion=None, tipo_map=None):
    """mapping: {norm(header): id_rex}. Suma columnas del mismo id. Cuadra al peso."""
    periodo = config["periodo"]; emp_id = config.get("empresa_id", ""); mut_id = config.get("mutual_id", "")
    apv_inst = config.get("apv_inst", "afp"); caja_inst = config.get("caja_inst", "losandes")
    ncont = config.get("num_contrato", 1); jornada = config.get("jornada", "C")
    # Tope imponible de salud = tope imponible AFP (mismo techo, ~90 UF). OJO: NO es topeSalud_pesos,
    # que es el 7% del tope (monto máx. de cotización), no la base imponible.
    tope_salud = _num(params_row.get("topeImp_pesos_afp", 0)); sis_pct = _num(params_row.get("sis", 0))
    if homolog:
        mut_id = resolver_inst(mut_id, homolog, {"mu"}) or mut_id
        caja_inst = resolver_inst(caja_inst, homolog, {"ca"}) or caja_inst
        apv_inst = resolver_inst(apv_inst, homolog, {"af"}) or apv_inst
    hdr = [x if str(x) != "nan" else "" for x in df.iloc[header_row].values]
    th_i = struct.get("total_haberes"); td_i = struct.get("total_descuentos")
    def sidx(k): return struct.get(k)
    _struct_idx = set(struct.values())   # columnas estructurales: NO son conceptos aunque su nombre coincida
    id_cols = OrderedDict()
    for i, h in enumerate(hdr):
        if i in _struct_idx: continue
        cid = mapping.get(norm(h))
        if cid: id_cols.setdefault(cid, []).append(i)
    def grp_of(cols):
        i = cols[0]
        if th_i is not None and i < th_i: return "haber"
        if td_i is not None and th_i is not None and th_i < i < td_i: return "desc"
        if td_i is not None and i > td_i: return "aporte"
        return "desc"
    filas = []; flags = set(); empleados = 0; omitidos = 0; bh = bd = bt = 0; log_contratos = []
    if sidx("dias_trab") is None:
        flags.add("El libro no trae 'Días Trabajados': se asumieron 30 días por trabajador — revisar con el consultor.")
    _ETIQ = {"af": "AFP", "is": "Salud", "mu": "Mutual", "ca": "Caja"}
    inst_seen = {}  # (clasif, valor_libro) -> id_rex resuelto (o None si no homologó)
    def _reg_inst(clasif, raw, resuelto):
        s = str(raw).strip()
        if not s or s.lower() in ("0", "nan", "none"): return
        inst_seen.setdefault((clasif, s), resuelto)
    for _, row in df.iloc[header_row+1:].iterrows():
        rut = str(row[sidx("rut")]).replace(".", "").strip().upper() if sidx("rut") is not None else ""
        if not rut or rut.lower() == "nan" or not rut[0].isdigit() or "total" in rut.lower(): continue
        # --- contrato/empresa/mutual por RUT (dotación) — primero, para poder OMITIR a quien no esté ---
        ncont_e, emp_e, mut_e, pmut_e, caja_e = ncont, emp_id, mut_id, None, caja_inst
        if dotacion:
            fecha_ing = row[sidx("fecha_ingreso")] if sidx("fecha_ingreso") is not None else None
            rc = resolver_contrato(rut, fecha_ing, periodo, dotacion)
            if not rc["ok"]:
                log_contratos.append({"rut": rut, "motivo": rc["motivo"]})
                omitidos += 1
                continue                                  # no está en la dotación -> se OMITE del archivo
            if rc["contrato"] not in (None, ""): ncont_e = rc["contrato"]
            emp_e = rc["empresa"] or emp_id
            mut_e = rc["mutual"] or mut_id
            pmut_e = rc["pmutual"]
            if rc.get("caja"): caja_e = rc["caja"]
            if homolog and mut_e:
                _mr = resolver_inst(mut_e, homolog, {"mu"}); _reg_inst("mu", mut_e, _mr); mut_e = _mr or mut_e
            if homolog and caja_e:
                _cr = resolver_inst(caja_e, homolog, {"ca"}); _reg_inst("ca", caja_e, _cr); caja_e = _cr or caja_e
        empleados += 1
        dt = int(_num(row[sidx("dias_trab")])) if sidx("dias_trab") is not None else 30
        afp_raw = (row[sidx("afp")] if sidx("afp") is not None and pd.notna(row[sidx("afp")]) else 0) or 0
        idafp_res = resolver_inst(afp_raw, homolog, {"af"}) if homolog else None
        idafp = idafp_res or afp_raw
        _reg_inst("af", afp_raw, idafp_res)
        sal = row[sidx("salud")] if sidx("salud") is not None and pd.notna(row[sidx("salud")]) else ""
        idsal_res = resolver_inst(sal, homolog, {"is"}) if homolog else None
        idsal = idsal_res or (norm(sal).replace("_", "").replace(" ", "") if sal else 0)
        _reg_inst("is", sal, idsal_res)
        cot_afp = _num(cot_hist.get(f"{periodo}{idafp}", 0)) * 100
        liq = _num(row[sidx("liquido")]) if sidx("liquido") is not None else 0
        sums = {cid: sum(_num(row[i]) for i in cols) for cid, cols in id_cols.items()}
        rebajas = sums.get("afp", 0) + sums.get("isapre", 0) + sums.get("cesEmpleado", 0)
        # Renta imponible AFP: del libro si viene; si NO (el libro no la trae), se DERIVA del monto de una
        # cotización que sea % puro del imponible (AFP ÷ tasa, o SIS ÷ tasa). No afecta la cuadratura
        # (que va por montos); solo alimenta la columna "Afecto". Se avisa para revisar.
        base_afp = _num(row[sidx("base_afp")]) if sidx("base_afp") is not None else 0
        if base_afp <= 0:
            afp_m = sums.get("afp", 0); sis_m = sums.get("sis", 0)
            if afp_m > 0 and cot_afp > 0:
                base_afp = round(afp_m / (cot_afp / 100.0))
                flags.add("El libro no trae la renta imponible: se derivó del AFP (monto ÷ tasa) — revisar con el consultor.")
            elif sis_m > 0 and sis_pct > 0:
                base_afp = round(sis_m / (sis_pct / 100.0))
                flags.add("El libro no trae la renta imponible: se derivó del SIS (monto ÷ tasa) — revisar con el consultor.")
        base_ces = _num(row[sidx("base_ces")]) if sidx("base_ces") is not None else base_afp
        if base_ces <= 0: base_ces = base_afp
        base_trib = _num(row[sidx("base_trib")]) if sidx("base_trib") is not None else 0
        pactado = _num(row[sidx("sueldo_pactado")]) if sidx("sueldo_pactado") is not None else 0
        emp_rows = []
        def add(cid, monto, afecto=0, inst=0, cot=0, init=0, reb=0, grp="desc", p7=0, p8=0):
            emp_rows.append((grp, [periodo, rut, ncont_e, cid, round(monto), round(afecto), inst, cot, 0, dt,
                             "x", emp_e, round(reb), 0, 0, jornada, "", round(init), 1, round(p7), round(p8)]))
        for cid, cols in id_cols.items():
            if cid == "impuesto": continue
            # Sobregiro: concepto dependiente de la POSICIÓN del mes (haber = compensaSobre, descuento =
            # sobregiro_anterior). Se re-decide por mes para soportar un mapeo único en modo multi-mes.
            if cid in ("compensaSobre", "sobregiro_anterior"):
                _gs = grp_of(cols)
                _cs = "compensaSobre" if _gs == "haber" else "sobregiro_anterior"
                add(_cs, sums[cid], grp=("haber" if _gs == "haber" else "desc")); continue
            m = sums[cid]; g = (tipo_map or {}).get(cid) or grp_of(cols)
            if cid == "afp": add(cid, m, afecto=base_afp, inst=idafp, cot=cot_afp, grp=g)
            elif cid == "isapre":
                af = base_afp if idsal == "fonasa" else (min(base_afp, tope_salud) if tope_salud else base_afp)
                # parcial8 (isapre) = tope imponible del mes
                add(cid, m, afecto=af, inst=idsal, grp=g, p8=tope_salud)
            elif cid == "cesEmpleado": add(cid, m, afecto=base_ces, inst=idafp, cot=0.6, grp=g)
            elif cid == "mutual": add(cid, m, afecto=base_afp, inst=mut_e, cot=(_num(pmut_e) if pmut_e not in (None, "") else 0), grp="aporte")
            elif cid == "sis": add(cid, m, afecto=base_afp, inst=idafp, cot=sis_pct, grp="aporte")
            # parcial8 (cesAporteSol) = imponible del mes (base del aporte solidario)
            elif cid in ("cesAporteCi", "cesAporteSol"):
                add(cid, m, afecto=base_ces, inst=idafp, grp="aporte", p8=(base_afp if cid == "cesAporteSol" else 0))
            elif cid == "aporteAFPemp": add(cid, m, afecto=base_afp, inst=idafp, cot=0.1, grp="aporte")
            elif cid == "aporteFAPPCEV": add(cid, m, afecto=base_afp, inst=idafp, grp="aporte")
            elif cid in INST_AFP_CONC: add(cid, m, inst=apv_inst, grp=g)
            elif cid in INST_CAJA_CONC: add(cid, m, inst=caja_e, grp=g)
            elif cid == "cajaComp": add(cid, m, afecto=base_afp, inst=caja_e, grp=g)
            elif cid in RELIQ:
                add(cid, m, afecto=base_afp, grp=("aporte" if cid in RELIQ_APO else "desc"))
                flags.add("Reliquidación cargada como consolidado — revisar devengo (Fecha de aplicación por mes de origen)")
            else:
                add(cid, m, init=((pactado or m) if cid == "sueldoBase" else 0), grp=g)
        trib = base_trib if base_trib else max(base_afp - rebajas, 0)
        # parcial7 (impuesto) = total de rebajas por leyes sociales (mismo valor que la col 12)
        add("impuesto", sums.get("impuesto", 0), afecto=trib, reb=rebajas, grp="desc", p7=rebajas)
        add("totalesEmpl", liq, afecto=base_afp, grp="total")
        H = sum(r[4] for g, r in emp_rows if g == "haber")
        D = sum(r[4] for g, r in emp_rows if g == "desc")
        T = next((r[4] for g, r in emp_rows if g == "total"), 0)
        TH = _num(row[sidx("total_haberes")]) if th_i is not None else H
        TD = _num(row[sidx("total_descuentos")]) if td_i is not None else D
        if abs(H - TH) > 2: bh += 1
        if abs(D - TD) > 2: bd += 1
        if abs((H - D) - liq) > 2: bt += 1   # líquido = haberes − descuentos (los aportes no cuentan)
        for g, r in emp_rows:
            if r[3] == "impuesto" or _num(r[4]) != 0 or r[3] == "totalesEmpl":
                filas.append(r)
    log_inst = [{"tipo": _ETIQ.get(cl, cl), "valor_libro": raw,
                 "id_rex": (res if res else ""), "estado": ("OK" if res else "SIN HOMOLOGAR")}
                for (cl, raw), res in sorted(inst_seen.items())]
    return filas, {"empleados": empleados, "omitidos": omitidos, "filas": len(filas), "flags": sorted(flags),
                   "descuadre_haberes": bh, "descuadre_descuentos": bd, "descuadre_liquido": bt,
                   "log_contratos": log_contratos, "log_inst": log_inst,
                   "homolog_vacia": not bool(homolog)}

def validar_cuadratura(df, header_row, struct, filas):
    """Compatibilidad: la cuadratura ahora viene en el resumen de generar_detalle."""
    return {"nota": "usar el resumen de generar_detalle"}
