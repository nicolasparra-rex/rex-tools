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
        try: fi = pd.to_datetime(row[cf]) if (cf < len(row) and not pd.isna(row[cf])) else None
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
"Rebaja por zona extrema","Jornada","Días de vacaciones","Monto Init","Fase"]

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

def classify_and_map(hdr, struct, catalog_names=None, saved=None):
    catalog_names = catalog_names or {}; saved = saved or {}
    th = struct.get("total_haberes"); td = struct.get("total_descuentos")
    struct_idx = set(struct.values())
    struct_syn = set(s for syns in STRUCT.values() for s in syns)
    filas = []
    for i, h in enumerate(hdr):
        n = norm(h)
        if not h or i in struct_idx or n in IGNORAR or n in struct_syn: continue
        if th is not None and i < th: grupo = "haber"
        elif td is not None and th is not None and th < i < td: grupo = "descuento"
        elif td is not None and i > td: grupo = "aporte"
        else: grupo = "?"
        cid, fuente, conf = None, "SIN MAPEAR", "-"
        if n in saved: cid, fuente, conf = saved[n], "guardado", "alta"
        elif n in CONCEPTO:
            cid = CONCEPTO[n]; fuente, conf = "diccionario", "alta"
            if cid == "__SOBREGIRO__": cid, conf = ("compensaSobre" if grupo == "haber" else "sobregiro_anterior"), "media"
            if cid == "__ISAPRE_AD__": cid = "isapre"
        elif n in catalog_names: cid, fuente, conf = catalog_names[n], "catalogo", "media"
        filas.append({"col": i+1, "header": h, "grupo": grupo, "id_rex": cid, "fuente": fuente, "confianza": conf})
    return filas

def generar_detalle(df, header_row, struct, mapping, params_row, cot_hist, config, homolog=None, dotacion=None):
    """mapping: {norm(header): id_rex}. Suma columnas del mismo id. Cuadra al peso."""
    periodo = config["periodo"]; emp_id = config.get("empresa_id", ""); mut_id = config.get("mutual_id", "")
    apv_inst = config.get("apv_inst", "afp"); caja_inst = config.get("caja_inst", "losandes")
    ncont = config.get("num_contrato", 1); jornada = config.get("jornada", "C")
    tope_salud = _num(params_row.get("topeSalud_pesos", 0)); sis_pct = _num(params_row.get("sis", 0))
    if homolog:
        mut_id = resolver_inst(mut_id, homolog, {"mu"}) or mut_id
        caja_inst = resolver_inst(caja_inst, homolog, {"ca"}) or caja_inst
        apv_inst = resolver_inst(apv_inst, homolog, {"af"}) or apv_inst
    hdr = [x if str(x) != "nan" else "" for x in df.iloc[header_row].values]
    th_i = struct.get("total_haberes"); td_i = struct.get("total_descuentos")
    def sidx(k): return struct.get(k)
    id_cols = OrderedDict()
    for i, h in enumerate(hdr):
        cid = mapping.get(norm(h))
        if cid: id_cols.setdefault(cid, []).append(i)
    def grp_of(cols):
        i = cols[0]
        if th_i is not None and i < th_i: return "haber"
        if td_i is not None and th_i is not None and th_i < i < td_i: return "desc"
        if td_i is not None and i > td_i: return "aporte"
        return "desc"
    filas = []; flags = set(); empleados = 0; bh = bd = bt = 0; log_contratos = []
    for _, row in df.iloc[header_row+1:].iterrows():
        rut = str(row[sidx("rut")]).replace(".", "").strip().upper() if sidx("rut") is not None else ""
        if not rut or rut.lower() == "nan" or not rut[0].isdigit() or "total" in rut.lower(): continue
        empleados += 1
        dt = int(_num(row[sidx("dias_trab")])) if sidx("dias_trab") is not None else 0
        base_afp = _num(row[sidx("base_afp")]) if sidx("base_afp") is not None else 0
        base_ces = _num(row[sidx("base_ces")]) if sidx("base_ces") is not None else base_afp
        base_trib = _num(row[sidx("base_trib")]) if sidx("base_trib") is not None else 0
        afp_raw = (row[sidx("afp")] if sidx("afp") is not None and pd.notna(row[sidx("afp")]) else 0) or 0
        idafp = (resolver_inst(afp_raw, homolog, {"af"}) if homolog else None) or afp_raw
        sal = row[sidx("salud")] if sidx("salud") is not None and pd.notna(row[sidx("salud")]) else ""
        idsal = (resolver_inst(sal, homolog, {"is"}) if homolog else None) or (norm(sal).replace("_", "").replace(" ", "") if sal else 0)
        cot_afp = _num(cot_hist.get(f"{periodo}{idafp}", 0)) * 100
        liq = _num(row[sidx("liquido")]) if sidx("liquido") is not None else 0
        sums = {cid: sum(_num(row[i]) for i in cols) for cid, cols in id_cols.items()}
        rebajas = sums.get("afp", 0) + sums.get("isapre", 0) + sums.get("cesEmpleado", 0)
        pactado = _num(row[sidx("sueldo_pactado")]) if sidx("sueldo_pactado") is not None else 0
        # --- resolución de contrato / empresa / mutual por RUT (dotación) ---
        ncont_e, emp_e, mut_e, pmut_e, caja_e = ncont, emp_id, mut_id, None, caja_inst
        if dotacion is not None:
            fecha_ing = row[sidx("fecha_ingreso")] if sidx("fecha_ingreso") is not None else None
            rc = resolver_contrato(rut, fecha_ing, periodo, dotacion)
            if rc["contrato"] not in (None, ""): ncont_e = rc["contrato"]
            emp_e = rc["empresa"] or emp_id
            mut_e = rc["mutual"] or mut_id
            pmut_e = rc["pmutual"]
            if rc.get("caja"): caja_e = rc["caja"]
            if homolog and mut_e: mut_e = resolver_inst(mut_e, homolog, {"mu"}) or mut_e
            if homolog and caja_e: caja_e = resolver_inst(caja_e, homolog, {"ca"}) or caja_e
            if not rc["ok"]: log_contratos.append({"rut": rut, "motivo": rc["motivo"]})
        emp_rows = []
        def add(cid, monto, afecto=0, inst=0, cot=0, init=0, reb=0, grp="desc"):
            emp_rows.append((grp, [periodo, rut, ncont_e, cid, round(monto), round(afecto), inst, cot, 0, dt,
                             "x", emp_e, round(reb), 0, 0, jornada, "", round(init), 1]))
        for cid, cols in id_cols.items():
            if cid == "impuesto": continue
            m = sums[cid]; g = grp_of(cols)
            if cid == "afp": add(cid, m, afecto=base_afp, inst=idafp, cot=cot_afp, grp=g)
            elif cid == "isapre":
                af = base_afp if idsal == "fonasa" else (min(base_afp, tope_salud) if tope_salud else base_afp)
                add(cid, m, afecto=af, inst=idsal, grp=g)
            elif cid == "cesEmpleado": add(cid, m, afecto=base_ces, inst=idafp, cot=0.6, grp=g)
            elif cid == "mutual": add(cid, m, afecto=base_afp, inst=mut_e, cot=(_num(pmut_e) if pmut_e not in (None, "") else 0), grp="aporte")
            elif cid == "sis": add(cid, m, afecto=base_afp, inst=idafp, cot=sis_pct, grp="aporte")
            elif cid in ("cesAporteCi", "cesAporteSol"): add(cid, m, afecto=base_ces, inst=idafp, grp="aporte")
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
        add("impuesto", sums.get("impuesto", 0), afecto=trib, reb=rebajas, grp="desc")
        add("totalesEmpl", liq, afecto=base_afp, grp="total")
        H = sum(r[4] for g, r in emp_rows if g == "haber")
        D = sum(r[4] for g, r in emp_rows if g == "desc")
        T = next((r[4] for g, r in emp_rows if g == "total"), 0)
        TH = _num(row[sidx("total_haberes")]) if th_i is not None else H
        TD = _num(row[sidx("total_descuentos")]) if td_i is not None else D
        if abs(H - TH) > 2: bh += 1
        if abs(D - TD) > 2: bd += 1
        if abs(T - liq) > 2: bt += 1
        for g, r in emp_rows:
            if r[3] == "impuesto" or _num(r[4]) != 0 or r[3] == "totalesEmpl":
                filas.append(r)
    return filas, {"empleados": empleados, "filas": len(filas), "flags": sorted(flags),
                   "descuadre_haberes": bh, "descuadre_descuentos": bd, "descuadre_liquido": bt,
                   "log_contratos": log_contratos}

def validar_cuadratura(df, header_row, struct, filas):
    """Compatibilidad: la cuadratura ahora viene en el resumen de generar_detalle."""
    return {"nota": "usar el resumen de generar_detalle"}
