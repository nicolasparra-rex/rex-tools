# -*- coding: utf-8 -*-
"""Motor de lectura de 'libro de remuneraciones' de cualquier cliente -> mapeo a conceptos Rex.
Enfoque: autodeteccion de estructura + clasificacion por bloque (posicion) + propuesta de ID
por sinonimos y catalogo del cliente. Lo no resuelto se marca para confirmacion humana."""
import unicodedata, pandas as pd
 
def norm(s):
    if s is None: return ""
    return "".join(c for c in unicodedata.normalize("NFD", str(s).strip().lower()) if unicodedata.category(c) != "Mn")
 
# ---------- Sinonimos de columnas ESTRUCTURALES (no son conceptos) ----------
STRUCT = {
 "rut": ["numero de documento","rut","n de documento","rut trabajador"],
 "nombre": ["nombre completo","nombre","apellido y nombre","nombre trabajador"],
 "dias_trab": ["dias trabajados"],
 "afp": ["fondo de cotizacion","afp","prevision"],
 "salud": ["fonasa/isapre","salud","isapre"],
 "base_afp": ["base imponible afp","imponible afp","imp. prev./salud","imponible topeado"],
 "base_ces": ["base imponible cesantia","imponible cesantia","imp. cesantia"],
 "base_trib": ["base tributable","tributable"],
 "total_haberes": ["total haberes"],
 "total_descuentos": ["total descuentos"],
 "total_aportes": ["total aportes"],
 "liquido": ["liquido a pago","liquido a recibir","liquido"],
 "plan_uf": ["plan isapre uf"],
 "tramo": ["tramo"],
 "centro_costo": ["centro de costo","id centro de costo","nombre c. costo"],
 "cargo": ["cargo","workday - grade","categoria"],
 "fecha_ingreso": ["fecha ingreso compania","fecha de alta","inicio contrato"],
 "id_ext": ["workday - id","empleado","id empleado"],
}
# columnas a ignorar del todo (metadatos)
IGNORAR = {"sueldo","plan uf","fecha de baja","departamento","tipo de empleado","id centro de costo","nombre c. costo","area","sede","workday - grade","grade","periodo","empleado"}
 
# ---------- Diccionario de sinonimos de CONCEPTOS (nombre -> id Rex) ----------
CONCEPTO = {
 "sueldo base":"sueldoBase","sueldo":"sueldoBase",
 "gratificacion":"gratificacion","diferencia de gratificacion":"gratificacion","diferencia gratificacion":"gratificacion",
 "bono ayuda hijo menor":"BonoHijoMenor","bono internet celular":"BonoCelular","bono sala cuna":"BonoSalaCunaG",
 "dif subsidio licencia medica":"DifSubLicenciaMedica","asignacion de alimentacion":"AsigAlimentacion",
 "asignacion de movilizacion":"movilizacion","movilizacion":"movilizacion","asignacion de teletrabajo":"AsignacionTeletrabjo",
 "diferencia a pagar":"DiferenciaAPagar","guardia pasiva":"GuardiaPasiva","hora extra guardia activa":"HEGuardiaActiva",
 "horas extras al 50 %":"horasEx50","horas extras al 50%":"horasEx50",
 "indemnizacion sustitutiva previo aviso":"iasMes","indemnizacion por vacaciones":"iasVacaciones",
 "bono extraordinario":"BonoExtraordinario","bono de demanda":"BonoDemanda","bono demanda":"Bono_Demanda",
 "bono extraordinario marketing":"BonoMarketing","aguinaldo":"AguinaldoMi","bono":"BonoMi",
 "sobregiro":"__SOBREGIRO__",  # resuelto por posicion (haber=compensaSobre / desc=sobregiro_anterior)
 # descuentos
 "descuento afp":"afp","cotiz. previ. obligatoria":"afp","cotizacion obligatoria previsional":"afp","a.f.p.":"afp",
 "cotizacion salud":"isapre","cotiz. salud obligatoria":"isapre","salud":"isapre",
 "cotizacion adicional isapre":"__ISAPRE_AD__","adicional salud":"__ISAPRE_AD__","adicional isapre":"__ISAPRE_AD__",
 "seguro de desempleo":"cesEmpleado","seguro cesantia":"cesEmpleado","seguro de cesantia":"cesEmpleado",
 "impuesto unico":"impuesto","impuesto":"impuesto",
 "cotiz. prev. voluntaria":"apvi","a.p.v.":"apvi","apv":"apvi",
 "anticipo sueldo":"anticipo","anticipo de sueldo":"anticipo","anticipo finiquito":"AnticipoPrestamoMi",
 "credito personal caja los andes":"cajaCred","leasing (ahorro) caja los andes":"cajaLeas",
 "seguro de vida caja los andes":"cajaVida","cuenta 2 pesos":"afpAhor","cuota sindical":"CuotaSindical",
 "descuento falp":"DescuentoFalp","descuento techops":"DescuentoTechops","retencion judicial":"retencionJudicial",
 "diferencia a descontar mes anterior":"sobregiro_anterior",
 # aportes empleador
 "aporte mutual":"mutual","mutual empleador":"mutual","aporte sis":"sis","sis":"sis",
 "cesantia empleador":"cesAporteCi","seguro cesantia empleador":"cesAporteCi",
 "ap. emp. cap. individual":"aporteAFPemp","afp prevision empleador":"aporteAFPemp",
 "ap. emp. ss exp. de vida":"aporteFAPPCEV","cotizacion expectativa de vida":"aporteFAPPCEV",
 # reliquidaciones (consolidado)
 "cotiz. previ. por reliquidacion":"reliquidaAfp","cotiz. salud obligatoria por reliquidacion":"reliquidaIsapre",
 "seguro cesantia por reliquidacion":"reliquidaCesEmpl","impuesto por reliquidacion":"reliquidaImpuesto",
 "sis por reliquidacion":"reliquidaSis","cesantia (empleador) por reliquidacion":"reliquidaCesCi",
 "mutual por reliquidacion":"reliquidaMutual","afp prevision empleador por reliquidacion":"reliquidaAporteAFP",
 "cotizacion expectativa de vida por reliquidacion":"reliquidaAporteCEV",
}
 
def load_grid(path, sheet_hint="libro"):
    xls = pd.ExcelFile(path)
    sh = [s for s in xls.sheet_names if sheet_hint in s.lower()]
    sh = sh[0] if sh else xls.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sh, header=None)
    return df, sh
 
def detect_header_row(df, key="numero de documento", alt="rut"):
    for r in range(min(15, len(df))):
        vals = [norm(x) for x in df.iloc[r].values]
        if any(key in v for v in vals) or any(v==alt for v in vals):
            return r
    return 0
 
def match_struct(hdr):
    """Devuelve {campo: indice} para columnas estructurales."""
    out = {}
    for i, h in enumerate(hdr):
        n = norm(h)
        if not n: continue
        for campo, syns in STRUCT.items():
            if campo in out: continue
            if any(n == s or n.startswith(s) for s in syns):
                out[campo] = i; break
    return out
 
def classify_and_map(hdr, struct, catalog_names=None, saved=None):
    """Para cada columna de concepto: grupo por posicion + id propuesto + fuente + confianza."""
    catalog_names = catalog_names or {}     # norm(nombre) -> id
    saved = saved or {}                      # norm(header) -> id
    th = struct.get("total_haberes"); td = struct.get("total_descuentos")
    struct_idx = set(struct.values())
    struct_syn = set(sy for syns in STRUCT.values() for sy in syns)
    filas = []
    for i, h in enumerate(hdr):
        n = norm(h)
        if not h or i in struct_idx or n in IGNORAR or n in struct_syn: continue
        # grupo por posicion
        if th is not None and i < th: grupo = "haber"
        elif td is not None and th is not None and th < i < td: grupo = "descuento"
        elif td is not None and i > td: grupo = "aporte"
        else: grupo = "?"
        # id propuesto
        cid, fuente, conf = None, "", ""
        if n in saved: cid, fuente, conf = saved[n], "guardado", "alta"
        elif n in CONCEPTO:
            cid = CONCEPTO[n]; fuente, conf = "diccionario", "alta"
            if cid == "__SOBREGIRO__": cid = "compensaSobre" if grupo=="haber" else "sobregiro_anterior"; conf="media"
            if cid == "__ISAPRE_AD__": cid = "isapre"; conf="alta"  # se fusiona con salud
        elif n in catalog_names: cid, fuente, conf = catalog_names[n], "catalogo", "media"
        else: fuente, conf = "SIN MAPEAR", "-"
        filas.append({"col": i+1, "header": h, "grupo": grupo, "id_rex": cid, "fuente": fuente, "confianza": conf})
    return filas
 
# ======================= GENERACIÓN DE MIGRACIÓN DETALLE =======================
OUT_COLS = ["Fecha de proceso","Id empleado","Número de contrato","Id del concepto",
"Monto del concepto","Afecto","Id de institución","Cotización de jubilación","Días de licencias",
"Días trabajados","Fecha de aplicación","Empresa","Total de rebajas por LLSS","Rentas no gravadas",
"Rebaja por zona extrema","Jornada","Días de vacaciones","Monto Init","Fase"]
 
# ids que reciben tratamiento especial (afecto/inst/cot)
LEG_DESC = {"afp","isapre","cesEmpleado","impuesto"}
APORTES  = {"mutual","sis","cesAporteCi","cesAporteSol","aporteAFPemp","aporteFAPPCEV"}
RELIQ    = {"reliquidaAfp","reliquidaIsapre","reliquidaCesEmpl","reliquidaImpuesto","reliquidaSis",
            "reliquidaMutual","reliquidaCesCi","reliquidaCesSol","reliquidaAporteAFP","reliquidaAporteCEV"}
INST_AFP_CONC = {"apvi","apvc","apviConvenido","afpAhor"}
INST_CAJA_CONC = {"cajaCred","cajaLeas","cajaVida"}
 
def _num(v):
    return v if isinstance(v,(int,float)) and pd.notna(v) else 0
 
def generar_detalle(df, header_row, struct, mapping, params_row, cot_hist, config):
    """mapping: {norm(header): id_rex}. config: dict con empresa_id, mutual_id, apv_inst,
    caja_inst, num_contrato, periodo, jornada. Devuelve (filas, resumen)."""
    periodo = config["periodo"]; emp_id = config.get("empresa_id",""); mut_id = config.get("mutual_id","")
    apv_inst = config.get("apv_inst","afp"); caja_inst = config.get("caja_inst","losandes")
    ncont = config.get("num_contrato",1); jornada = config.get("jornada","C")
    tope_salud = _num(params_row.get("topeSalud_pesos",0)); sis_pct = _num(params_row.get("sis",0))
    hdr = [x if str(x)!="nan" else "" for x in df.iloc[header_row].values]
    th_i = struct.get("total_haberes"); td_i = struct.get("total_descuentos")
    def sidx(k): return struct.get(k)
    # columnas por id
    isapre_cols = [i for i,h in enumerate(hdr) if mapping.get(norm(h))=="isapre"]
    salud_base_cols = isapre_cols  # salud + adicional van a isapre
    filas=[]; flags=set()
    data = df.iloc[header_row+1:]
    empleados=0
    for _,row in data.iterrows():
        rut = str(row[sidx("rut")]).replace(".","").strip() if sidx("rut") is not None else ""
        if not rut or rut.lower()=="nan" or not rut[0].isdigit() or "total" in rut.lower(): continue
        empleados+=1
        dt = int(_num(row[sidx("dias_trab")])) if sidx("dias_trab") is not None else 0
        base_afp = _num(row[sidx("base_afp")]) if sidx("base_afp") is not None else 0
        base_ces = _num(row[sidx("base_ces")]) if sidx("base_ces") is not None else base_afp
        base_trib = _num(row[sidx("base_trib")]) if sidx("base_trib") is not None else 0
        idafp = row[sidx("afp")] if sidx("afp") is not None and pd.notna(row[sidx("afp")]) else 0
        idafp = idafp or 0
        sal = row[sidx("salud")] if sidx("salud") is not None and pd.notna(row[sidx("salud")]) else ""
        idsal = norm(sal).replace("_","").replace(" ","") if sal else 0
        cot_afp = _num(cot_hist.get(f"{periodo}{idafp}",0))*100
        liq = _num(row[sidx("liquido")]) if sidx("liquido") is not None else 0
        # rebajas LLSS = afp + salud(isapre) + cesEmpleado
        rebajas = 0
        for i,h in enumerate(hdr):
            cid = mapping.get(norm(h))
            if cid in ("afp","isapre","cesEmpleado"): rebajas += _num(row[i])
        def add(cid,monto,afecto=0,inst=0,cot=0,init=0,reb=0,rent=0,aplic="x"):
            filas.append([periodo,rut,ncont,cid,round(monto),round(afecto),inst,cot,0,dt,aplic,
                          emp_id,round(reb),round(rent),0,jornada,"",round(init),1])
        isapre_done=False; impuesto_done=False
        for i,h in enumerate(hdr):
            cid = mapping.get(norm(h))
            if not cid: continue
            m = _num(row[i])
            if cid=="isapre":
                if not isapre_done:
                    tot = sum(_num(row[j]) for j in salud_base_cols)
                    if tot:
                        af = base_afp if idsal=="fonasa" else (min(base_afp,tope_salud) if tope_salud else base_afp)
                        add("isapre",tot,afecto=af,inst=idsal)
                    isapre_done=True
                continue
            if cid=="impuesto":
                trib = base_trib if base_trib else max(base_afp-rebajas,0)
                add("impuesto",m,afecto=trib,reb=rebajas); impuesto_done=True; continue
            if not m: continue
            if cid=="afp": add("afp",m,afecto=base_afp,inst=idafp,cot=cot_afp)
            elif cid=="cesEmpleado": add("cesEmpleado",m,afecto=base_ces,inst=idafp,cot=0.6)
            elif cid=="mutual": add("mutual",m,afecto=base_afp,inst=mut_id)
            elif cid=="sis": add("sis",m,afecto=base_afp,inst=idafp,cot=sis_pct)
            elif cid in ("cesAporteCi","cesAporteSol"): add(cid,m,afecto=base_ces,inst=idafp)
            elif cid=="aporteAFPemp": add(cid,m,afecto=base_afp,inst=idafp,cot=0.1)
            elif cid=="aporteFAPPCEV": add(cid,m,afecto=base_afp,inst=idafp)
            elif cid in INST_AFP_CONC: add(cid,m,inst=apv_inst)
            elif cid in INST_CAJA_CONC: add(cid,m,inst=caja_inst)
            elif cid in RELIQ:
                add(cid,m,afecto=base_afp); flags.add(f"Reliquidación {cid}: revisar devengo (Fecha de aplicación por mes de origen)")
            else:
                init = m if cid=="sueldoBase" else 0
                add(cid,m,init=init)
        if not impuesto_done:
            trib = base_trib if base_trib else max(base_afp-rebajas,0)
            add("impuesto",0,afecto=trib,reb=rebajas)
        add("totalesEmpl",liq,afecto=base_afp)
    return filas, {"empleados":empleados,"flags":sorted(flags)}
 
def validar_cuadratura(df, header_row, struct, filas):
    """Compara suma de haberes/descuentos/totalesEmpl del output contra el libro por RUT."""
    from collections import defaultdict
    hdr=[x if str(x)!="nan" else "" for x in df.iloc[header_row].values]
    th_i=struct.get("total_haberes"); td_i=struct.get("total_descuentos")
    def sidx(k): return struct.get(k)
    src={}
    for _,row in df.iloc[header_row+1:].iterrows():
        rut=str(row[sidx("rut")]).replace(".","").strip() if sidx("rut") is not None else ""
        if not rut or rut.lower()=="nan" or not rut[0].isdigit(): continue
        src[rut]=(_num(row[sidx("total_haberes")]) if th_i is not None else 0,
                  _num(row[sidx("total_descuentos")]) if td_i is not None else 0,
                  _num(row[sidx("liquido")]) if sidx("liquido") is not None else 0)
    APO=APORTES|{"reliquidaSis","reliquidaMutual","reliquidaCesCi","reliquidaCesSol","reliquidaAporteAFP","reliquidaAporteCEV"}
    DES={"afp","isapre","cesEmpleado","impuesto","apvi","apvc","apviConvenido","afpAhor","cajaCred","cajaLeas","cajaVida",
         "CuotaSindical","DescuentoFalp","DescuentoTechops","AnticipoPrestamoMi","anticipo","retencionJudicial",
         "sobregiro_anterior","reliquidaAfp","reliquidaIsapre","reliquidaCesEmpl","reliquidaImpuesto"}
    hab=defaultdict(float);des=defaultdict(float);tot={}
    for f in filas:
        rut=f[1];cid=f[3];m=f[5-1]  # monto col index 4
        m=f[4]
        if cid=="totalesEmpl": tot[rut]=m
        elif cid in APO: continue
        elif cid in DES: des[rut]+=m
        else: hab[rut]+=m
    bh=bd=bt=0
    for k,(TH,TD,LQ) in src.items():
        if abs(hab[k]-TH)>2: bh+=1
        if abs(des[k]-TD)>2: bd+=1
        if abs(tot.get(k,0)-LQ)>2: bt+=1
    return {"empleados":len(src),"descuadre_haberes":bh,"descuadre_descuentos":bd,"descuadre_liquido":bt}