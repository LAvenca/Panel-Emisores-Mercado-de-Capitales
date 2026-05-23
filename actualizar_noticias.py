import urllib.request
import urllib.parse
import re
import ssl
import html
import html as html_mod
import zipfile
import json
from datetime import datetime, timedelta

# ============================================================
# 1. LECTURA NATIVA DEL EXCEL
# ============================================================
def leer_emisores_excel(ruta):
    emisores = {}
    try:
        with zipfile.ZipFile(ruta, 'r') as z:
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                ss_xml = z.read('xl/sharedStrings.xml').decode('utf-8')
                shared_strings = re.findall(r'<t[^>]*>(.*?)</t>', ss_xml)
            xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
            filas = re.findall(r'<row[^>]*r="(\d+)"[^>]*>(.*?)</row>', xml, re.DOTALL)
            for r_num, row_content in filas:
                if int(r_num) < 2:
                    continue
                celdas = re.findall(r'<c [^>]*>(.*?)</c>', row_content, re.DOTALL)
                valores = []
                for celda_full in celdas:
                    v = re.search(r'<v>(.*?)</v>', celda_full)
                    if not v:
                        valores.append("")
                        continue
                    raw = v.group(1).strip()
                    try:
                        idx = int(raw)
                        if idx < len(shared_strings):
                            valores.append(re.sub(r'<[^>]+>', '', shared_strings[idx]).strip())
                        else:
                            valores.append(raw)
                    except ValueError:
                        valores.append(raw)
                if len(valores) >= 2 and valores[0] and valores[1]:
                    emisores[valores[0]] = valores[1]
    except Exception as e:
        print(f"Error leyendo Excel: {e}")
    return emisores

# ============================================================
# 2. SEMÁFORO DE RIESGO
# ============================================================
SEMAFORO = {
    "RUTAS DE LIMA S.A.":                   "rojo",
    "AUNA S.A.":                            "amarillo",
    "VOLCAN COMPAÑIA MINERA S.A.A.":        "rojo",
    "ALICORP S.A.A.":                       "verde",
    "CREDICORP LTD.":                       "verde",
    "BANCO DE CREDITO DEL PERU":            "verde",
    "INTERBANK":                            "verde",
    "BBVA PERU":                            "verde",
    "BBVA MEXICO":                          "amarillo",
    "SCOTIABANK PERU S.A.A.":              "verde",
    "MIBANCO":                              "verde",
    "BANCO GNB PERU S.A.":                  "amarillo",
    "FINANCIERA EFECTIVA S.A.":             "amarillo",
    "BANCO EFECTIVA":                       "amarillo",
    "CERRO VERDE":                          "verde",
    "SOUTHERN PERU COPPER CORPORATION":     "verde",
    "ACEROS AREQUIPA":                      "verde",
    "INRETAIL PERU CORP.":                  "verde",
    "ADMINISTRADORA JOCKEY PLAZA S.A.":     "amarillo",
    "BANCO DE LA NACION":                   "verde",
}

def get_semaforo(emisor):
    nombre = limpiar_nombre(emisor).upper()
    if nombre in SEMAFORO:
        return SEMAFORO[nombre]
    for clave, color in SEMAFORO.items():
        if clave.upper() in nombre or nombre in clave.upper():
            return color
    return None

def semaforo_html(color):
    if not color:
        return ""
    cfg = {
        "rojo":     ("🔴", "bg-red-500/20 border-red-500/50 text-red-300",            "Riesgo Alto"),
        "amarillo": ("🟡", "bg-yellow-500/20 border-yellow-500/50 text-yellow-300",   "Riesgo Moderado"),
        "verde":    ("🟢", "bg-emerald-500/20 border-emerald-500/50 text-emerald-300","Riesgo Bajo"),
    }
    if color not in cfg:
        return ""
    icono, clases, label = cfg[color]
    return f'<span class="text-[9px] border px-1.5 py-0.5 rounded font-bold {clases}">{icono} {label}</span>'

# ============================================================
# 3. SCORING SEGÚN CRITERIOS DE PRIORIZACIÓN OFICIALES
# ============================================================
EMISORES_PRIORITARIOS = ["rutas de lima", "auna"]
SCORE_MINIMO_RESUMEN  = 5

# NIVEL 10 — Alta Prioridad Absoluta: Downgrade / Default / Estrés financiero
# Aplica a Renta Fija, Variable y Fondos
SCORING_NOTICIAS = [
    (10, [
        # Downgrade / rebaja de calificación
        'downgrade','rebaja','rebaja de calificacion','rebaja calificacion',
        'rebaja de nota','rebaja crediticia','rebaja de rating',
        'reduccion de calificacion','reducción de calificación',
        'baja de nota','baja la calificacion','baja la nota',
        'degradacion crediticia','coloca en revision a la baja',
        'revision a la baja','watch negative','creditwatch negative',
        'on review for downgrade','rebaja deuda','rebaja bonos',
        'rebaja notas de deuda','rebaja deuda de largo plazo',
        # Default / estrés financiero (RF)
        'default','incumplimiento de pago','incumple','suspende pagos',
        'quiebra','bancarrota','concurso de acreedores',
        'reestructuracion de deuda','reestructuración de deuda',
        'refinanciamiento forzoso','moratoria','cesacion de pagos',
        'deterioro de liquidez','deterioro de solvencia',
        'crisis financiera','estrés financiero',
        # Fondos: restricción de rescates / liquidación
        'restriccion de rescates','restricción de rescates',
        'suspende rescates','liquidacion del fondo','liquidación del fondo',
        'suspende operaciones','cierre del fondo',
        'tracking error elevado','ciberataque',
    ]),

    (9, [
        # Alta Prioridad Absoluta: Fusiones / Adquisiciones / Cambios de control
        'fusion','fusión','adquisicion','adquisición','compra de','adquiere',
        'absorcion','absorción','cambio de control','toma de control',
        'oferta publica de adquisicion','opa','reorganizacion corporativa',
        'reorganización corporativa','escision','escisión','spin-off',
        'joint venture','alianza estrategica','alianza estratégica',
        # RV: cambios relevantes en directorio / management / CIO
        'renuncia del ceo','cambio de ceo','nuevo ceo',
        'sale el director','cambio en directorio','nuevo directorio',
        'renuncia del cfo','nuevo cfo','cambio de gerente general',
        'sale portfolio manager','nuevo cio','renuncia cio',
        'cambio en gobierno corporativo',
    ]),

    (8, [
        # Upgrade / mejora de calificación (Alta Prioridad RF)
        'upgrade','mejora de calificacion','mejora calificacion',
        'mejora de nota','mejora de rating','eleva calificacion',
        'eleva la nota','eleva la calificacion','sube calificacion',
        'sube la nota','alza de calificacion','alza de nota',
        'revision al alza','watch positive','creditwatch positive',
        'on review for upgrade','mejora deuda','mejora bonos',
        'mejora deuda de largo plazo',
        # Outlook negativo (RF/RV)
        'outlook negativo','perspectiva negativa',
        'cambia perspectiva a negativa','cambia outlook a negativo',
        'perspectiva negativa deuda','perspectiva negativa bonos',
        'negative watch','under review negative','en revision negativa',
        # Resultados fuera de expectativas (RV/RF)
        'por debajo de expectativas','supera expectativas',
        'resultado negativo sorpresa','perdida inesperada',
        'profit warning','guidance reducido','guidance rebajado',
        'resultados decepcionantes','resultados superan',
        # Multas / litigios / sanciones (RF/RV/Fondos)
        'multa','sancion','sanción','investigacion regulatoria',
        'investigación regulatoria','fraude','corrupcion','corrupción',
        'manipulacion de mercado','manipulación de mercado',
        'conflicto de interes','conflicto de interés',
        'escandalo','escándalo','denuncia penal',
        'indecopi','sbs sanciona','smv sanciona',
        'proceso judicial','sentencia condenatoria',
    ]),

    (7, [
        # Outlook positivo (RF)
        'outlook positivo','perspectiva positiva',
        'cambia perspectiva a positiva','cambia outlook a positivo',
        'perspectiva positiva deuda','perspectiva positiva bonos',
        'positive watch','under review positive','en revision positiva',
        # Clasificadoras (RF alta prioridad)
        "moody's",'moody','fitch','s&p','standard & poor',
        'japan credit rating','pacific credit rating',
        "moody's local",'apoyo & asociados','apoyo y asociados',
        'classifica','clasifica','califica','emite opinion',
        'informe de clasificacion','informe de calificacion',
        # Nuevas emisiones de deuda / aumentos de capital (RF/RV media)
        'nueva emision','nueva emisión','emite bonos','coloca bonos',
        'emision de deuda','emisión de deuda','aumento de capital',
        'oferta publica de acciones','ipo','suscripcion de acciones',
        'recompra de acciones','buyback','dividendo extraordinario',
        'cambio en dividendos','suspension de dividendos',
    ]),

    (6, [
        # Outlook estable / cambio de perspectiva general (RF)
        'outlook estable','perspectiva estable',
        'cambia perspectiva','cambia outlook','modifica perspectiva',
        'creditwatch','bajo revision','bajo revisión',
        'en revision','en revisión',
        'afirma calificacion','confirma calificacion',
        'ratifica calificacion','mantiene calificacion',
        # Regulatorio con impacto material (RF/RV/Fondos)
        'cambio regulatorio','nueva regulacion','nueva regulación',
        'basilea','limite regulatorio','límite regulatorio',
        'nueva norma sbs','nueva norma smv','nueva norma bcr',
        'requerimiento de capital','provisiones obligatorias',
        # Fondos: cambios en AUM / flujos / estrategia
        'salida de flujos','retiro masivo','flujos negativos',
        'caida de aum','caída de aum','reduccion de aum',
        'cambio de estrategia','cambio de benchmark',
        'cambio de politica de inversion','cambio de política de inversión',
        'cambio de custodio','fusiona fondos','fusión de fondos',
    ]),

    (5, [
        # Calificación crediticia general
        'calificacion','calificación','clasificacion','clasificación',
        'rating','nota crediticia','investment grade','grado de inversion',
        'grado de inversión','speculative grade','grado especulativo',
        'deuda de largo plazo','deuda largo plazo','bonos de largo plazo',
        'notas de deuda','calificacion de bonos','calificacion de deuda',
        # Hecho de importancia
        'hecho de importancia','hecho relevante','material fact',
        # Eventos macro / sectoriales (media prioridad RF/RV)
        'riesgo soberano','riesgo pais','riesgo país',
        'spread soberano','spread credito','credit spread',
        'evento geopolitico','evento geopolítico',
    ]),

    (4, [
        # CAPEX / inversiones relevantes (RF/RV media)
        'capex','inversion relevante','inversión relevante',
        'proyecto de expansion','proyecto de expansión',
        'nueva planta','nueva infraestructura',
        'contrato relevante','concesion relevante','concesión relevante',
        'adjudicacion relevante','adjudicación relevante',
        'licitacion ganada','licitación ganada',
        # Resultados financieros (RV/RF media)
        'ebitda','resultado operativo','resultado neto',
        'utilidad neta','perdida neta','pérdida neta',
        'ingresos trimestrales','resultado trimestral',
        'resultado semestral','resultado anual',
        'margen operativo','flujo de caja libre','free cash flow',
        # Fondos: variaciones AUM / cambios operativos
        'variacion de aum','variación de aum','nuevos flujos',
        'lanzamiento de fondo','nuevo etf','nuevo fondo',
        'cambio de comisiones','cambio de custodio',
    ]),

    (3, [
        # Financiamiento / deuda general
        'bonos','bono','deuda','deuda corporativa',
        'emision','emisión','sindicado','prestamo sindicado',
        'prestamo','préstamo','credito','crédito',
        'linea de credito','línea de crédito',
        'financiamiento','refinanciamiento',
        'colocacion','colocación','oferta publica',
        'spread','tasa de interes','tasa de interés',
        'cupón','cupon','vencimiento de deuda','amortizacion',
        # Estructura accionaria (RV)
        'cambio en estructura accionaria','nuevo accionista',
        'venta de participacion','venta de participación',
        'aumento de participacion','aumento de participación',
    ]),

    (2, [
        # Expansión / proyectos generales
        'proyecto','expansion','expansión','planta',
        'infraestructura','contrato','concesion','concesión',
        'adjudicacion','licitacion','obra','ampliacion','construccion',
        # RV: cambios en recomendaciones de analistas
        'recomendacion de analista','recomendación de analista',
        'precio objetivo','target price','eleva precio objetivo',
        'rebaja precio objetivo','inicia cobertura',
        'sobrepondera','infrapondera','mantener',
    ]),

    (1, [
        # Resultados genéricos / baja prioridad
        'utilidad','utilidades','ganancia','perdida','pérdida',
        'resultado','resultados','ingresos','revenue',
        'trimestre','semestre','anual','balance',
        'flujo de caja','margen','rentabilidad',
        'dividendo','dividendos','acciones','bolsa','bvl',
        # Fondos: actualizaciones rutinarias
        'actualizacion operativa','actualización operativa',
        'cambio menor','cambio organizacional',
    ]),
]

ALERTAS_REALES = [
    # Downgrade explícito
    'downgrade','rebaja la calificacion','rebaja calificacion',
    'rebaja de calificacion','rebaja la nota','baja la calificacion',
    'baja la nota','degrada calificacion','coloca en revision a la baja',
    'rebaja deuda de largo plazo','rebaja bonos','rebaja notas de deuda',
    # Default / estrés
    'default','incumplimiento de pago','suspende pagos','quiebra',
    'concurso de acreedores','reestructuracion de deuda',
    'restriccion de rescates','liquidacion del fondo',
    # Upgrade
    'upgrade','mejora la calificacion','mejora calificacion',
    'eleva la calificacion','eleva calificacion','sube la calificacion',
    'alza la calificacion','revision al alza',
    'mejora deuda de largo plazo','mejora bonos',
    # Outlook con dirección
    'outlook negativo','perspectiva negativa',
    'outlook positivo','perspectiva positiva',
    'cambia perspectiva a negativa','cambia perspectiva a positiva',
    'cambia outlook a negativo','cambia outlook a positivo',
    'watch negative','watch positive',
    'creditwatch negative','creditwatch positive',
    # Fusiones / adquisiciones
    'fusion','fusión','adquisicion','adquisición','cambio de control',
    'toma de control','opa',
    # Agencias con verbo
    "moody's rebaja","moody's eleva","moody's baja","moody's sube",
    "moody's coloca","moody's cambia","moody's afirma",
    "fitch rebaja","fitch eleva","fitch baja","fitch sube",
    "fitch coloca","fitch cambia","fitch afirma",
    "s&p rebaja","s&p eleva","s&p baja","s&p sube",
    "s&p coloca","s&p cambia","s&p afirma",
    "standard & poor rebaja","standard & poor eleva",
    "moody rebaja","moody eleva","moody baja","moody sube","moody coloca",
    "apoyo & asociados rebaja","apoyo & asociados eleva",
    "pacific credit rating rebaja","pacific credit rating eleva",
    "moody's local rebaja","moody's local eleva",
    # Multas / fraude
    'fraude','corrupcion','manipulacion de mercado',
    'sbs sanciona','smv sanciona','sentencia condenatoria',
    # Management
    'renuncia del ceo','sale el director','nuevo ceo',
    'sale portfolio manager','renuncia cio',
]

# Labels para el resumen visual
SCORE_LABELS = {
    10: ("🔴", "Downgrade / Default / Estrés"),
    9:  ("🟣", "Fusión / Adquisición / Control"),
    8:  ("🟠", "Upgrade / Outlook Neg. / Resultado / Multa"),
    7:  ("🔵", "Outlook Pos. / Clasificadora / Emisión"),
    6:  ("🟡", "Perspectiva / Regulatorio / Fondos AUM"),
    5:  ("⚪", "Calificación / Hecho Importancia / Macro"),
    4:  ("🏗",  "CAPEX / Resultados / Fondos Operativo"),
    3:  ("📄", "Deuda / Financiamiento / Accionaria"),
    2:  ("📌", "Analistas / Expansión / Proyectos"),
    1:  ("📊", "Resultados Generales / Baja Prioridad"),
}

def calcular_score(titulo):
    titulo_lower = titulo.lower()
    for score, keywords in SCORING_NOTICIAS:
        if any(k in titulo_lower for k in keywords):
            return score
    return 0

def es_alerta_real(titulo):
    titulo_lower = titulo.lower()
    return any(k in titulo_lower for k in ALERTAS_REALES)

def es_emisor_prioritario(emisor):
    return any(p in limpiar_nombre(emisor).lower() for p in EMISORES_PRIORITARIOS)

# ============================================================
# 4. CONFIGURACIÓN GLOBAL
# ============================================================
contexto = ssl._create_unverified_context()
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/122.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'es-419,es;q=0.9,en;q=0.8'
}

PALABRAS_CREDITICIAS = [
    'downgrade','upgrade','outlook','perspectiva','rating','calificacion',
    'calificación','moody','fitch','s&p','standard & poor','investment grade',
    'grado de inversion','grado de inversión','watch negative','watch positive',
    'creditwatch','bajo revision','bajo revisión','rebaja','mejora crediticia',
    'deuda','bonos','emision','emisión','sindicado','hecho de importancia',
    'aumento de capital','soberano','riesgo pais','riesgo país','spread',
    'deuda de largo plazo','notas de deuda','bonos de largo plazo',
    'fusion','fusión','adquisicion','adquisición','default',
    'pacific credit rating','apoyo & asociados','moody local',
]

PALABRAS_IGNORAR = {
    'de','la','el','los','las','del','y','en','al',
    'sa','saa','sac','sab','the','of','and'
}
TOKENS_GENERICOS = {
    'banco','fondo','corp','group','grupo','financiera',
    'inversiones','holding','capital','asset','trust','management'
}
PREFIJOS_IGNORAR = {
    'administradora','administrador','compania','compañia',
    'empresa','grupo','corporacion','corporación','sociedad'
}
EMISORES_EXTRANJEROS_KEYWORDS = [
    'luxembourg','ireland','ireland limited','international',
    'u.k.','europe','sicav','llp','asset management'
]
PALABRAS_SOBERANAS = [
    'gobierno','republica','republic','soberano','estado peruano',
    'gobierno peruano','ministerio','mef','bcr','bcrp'
]
TEMAS_FINANCIEROS_SOBERANOS = [
    'deuda','bonos','fiscal','presupuesto','deficit','déficit',
    'inversion','inversión','economia','economía','financiamiento',
    'rating','calificacion','calificación','moody','fitch','s&p',
    'downgrade','upgrade','outlook','perspectiva','credito','crédito',
    'gdp','pbi','pib','inflacion','inflación','reservas','mef',
    'banco central','tipo de cambio','tesoro','spread',
    'riesgo pais','riesgo país','impuesto','recaudacion','balanza',
    'deuda de largo plazo','notas de deuda',
]
TEMAS_BLACKLIST = [
    'narcotrafico','narcotráfico','droga','drogas','cocaina','cocaína',
    'crimen','sicario','homicidio','asesinato','feminicidio',
    'terremoto','sismo','huaico','inundacion','inundación',
    'futbol','fútbol','deporte','partido','concierto','festival',
    'musica','música','cantante','artista','actor','actriz',
    'farandula','farándula','espectaculo','espectáculo','television',
    'accidente','incendio','rescate','messi','neymar','ronaldo',
    'deportistas','multimillonario','celebridad','celebrity',
    'elecciones','candidato','keiko','fujimori','castillo','boluarte',
    'congreso','vacancia','violencia','delincuencia','extorsion',
    'maraton','maratón','ciclovia','ciclovía','avenida cerrada',
    'desvio','desvío','cierre vial','cierre temporal',
    'rutas nacionales','red vial','red federal','kilómetros de ruta',
    'licitacion de rutas','licitación de rutas','milei',
]
PAISES_EXTRANJEROS = [
    'argentina','colombi','bogota','bogotá','brasil',
    'venezuela','ecuador','bolivia','uruguay','paraguay'
]

ALIAS_EMISORES = {
    "jockey plaza":        "Jockey Plaza",
    "real plaza":          "Real Plaza",
    "open plaza":          "Open Plaza",
    "mall aventura":       "Mall Aventura",
    "saga falabella":      "Falabella",
    "tiendas peruanas":    "Oechsle",
    "inretail":            "InRetail",
    "intercorp":           "Intercorp",
    "credicorp":           "Credicorp",
    "prima afp":           "Prima AFP",
    "habitat afp":         "AFP Habitat",
    "profuturo":           "Profuturo",
    "cerro verde":         "Cerro Verde",
    "southern peru":       "Southern Peru Copper",
    "volcan":              "Volcan Compania Minera",
    "alicorp":             "Alicorp",
    "aceros arequipa":     "Aceros Arequipa",
    "banco de la nacion":  "Banco de la Nación Perú",
    "banco gnb":           "Banco GNB Perú",
    "mibanco":             "Mibanco",
    "financiera efectiva": "Banco Efectiva",
    "banco efectiva":      "Banco Efectiva",
    "bbva peru":           "BBVA Perú",
    "bbva mexico":         "BBVA México",
    "scotiabank":          "Scotiabank Perú",
    "bcp":                 "Banco de Crédito del Perú",
    "interbank":           "Interbank Perú",
    "brown brothers":      "Brown Brothers Harriman",
    "fibra prime":         "Fibra Prime",
    "fossal":              "Fossal",
    "laive":               "Laive",
    "cementos pacasmayo":  "Cementos Pacasmayo",
    "ferreyros":           "Ferreyros",
    "luz del sur":         "Luz del Sur",
    "pluz energia":        "Pluz Energía",
    "endispc":             "Enel Distribución Perú",
    "enel distribuc":      "Enel Distribución Perú",
    "enel":                "Enel Perú",
    "edelnor":             "Enel Distribución Perú",
    "telefonica":          "Telefónica del Perú",
    "entel":               "Entel Perú",
    "lima airport":        "Lima Airport Partners",
    "lap":                 "Lima Airport Partners",
    "rutas de lima":       "Rutas de Lima",
    "auna s.a":            "Auna",
    "auna oncologia":      "Auna",
    "auna salud":          "Auna",
    "gobierno peruano":    "Gobierno del Perú",
    "republica del peru":  "República del Perú",
}

LIMITE_FECHA  = datetime.now() - timedelta(days=14)
fecha_reporte = datetime.now().strftime("%d/%m/%Y %H:%M")
cache_prioritarias = []

# ============================================================
# 5. UTILIDADES DE FECHA
# ============================================================
def parsear_fecha(fecha_raw):
    formatos = [
        "%a, %d %b %Y %H:%M:%S %z",
        "%a, %d %b %Y %H:%M:%S GMT",
        "%a, %d %b %Y %H:%M:%S +0000",
    ]
    for fmt in formatos:
        try:
            return datetime.strptime(fecha_raw.strip(), fmt).replace(tzinfo=None)
        except Exception:
            pass
    m = re.search(r'(\d{1,2})\s+(\w{3})\s+(\d{4})', fecha_raw)
    if m:
        try:
            return datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%d %b %Y")
        except Exception:
            pass
    return None

def formatear_fecha(fecha_raw):
    dt = parsear_fecha(fecha_raw)
    if not dt:
        return fecha_raw[:11] if fecha_raw else "S/F"
    meses = {"Jan":"Ene","Feb":"Feb","Mar":"Mar","Apr":"Abr","May":"May",
              "Jun":"Jun","Jul":"Jul","Aug":"Ago","Sep":"Sep","Oct":"Oct",
              "Nov":"Nov","Dec":"Dic"}
    d = dt.strftime("%d %b %Y")
    for en, es in meses.items():
        d = d.replace(en, es)
    return d

def es_reciente(fecha_raw):
    dt = parsear_fecha(fecha_raw)
    if not dt:
        return True
    return dt >= LIMITE_FECHA

# ============================================================
# 6. UTILIDADES DE NOMBRE Y FILTROS
# ============================================================
def limpiar_nombre(emisor):
    nombre = html_mod.unescape(emisor)
    return re.sub(r'\s+', ' ', nombre).strip()

def es_emisor_extranjero(emisor):
    return any(kw in limpiar_nombre(emisor).lower() for kw in EMISORES_EXTRANJEROS_KEYWORDS)

def es_emisor_soberano(emisor):
    return any(p in limpiar_nombre(emisor).lower() for p in PALABRAS_SOBERANAS)

def titulo_es_relevante_financiero(titulo):
    return not any(t in titulo.lower() for t in TEMAS_BLACKLIST)

def titulo_relevante_para_emisor(titulo, emisor):
    titulo_lower = titulo.lower()
    nombre_lower = limpiar_nombre(emisor).lower()

    if not titulo_es_relevante_financiero(titulo):
        return False

    if 'bbva' in nombre_lower and 'mexico' in nombre_lower:
        CONTEXTO_BBVA_MX = [
            'bbva','mexico','méxico','banamex','moody','fitch','s&p',
            'downgrade','upgrade','calificacion','rating','perspectiva',
            'outlook','deuda','bonos','soberano','banco','rebaja','eleva'
        ]
        return any(c in titulo_lower for c in CONTEXTO_BBVA_MX)

    if 'efectiva' in nombre_lower or 'banco efectiva' in nombre_lower:
        CONTEXTO_EFECTIVA = [
            'efectiva','banco efectiva','financiera efectiva',
            'sbs','conversion','conversión','bancario','autoriza',
            'calificacion','rating','deuda','bonos','resultado','banco'
        ]
        return any(c in titulo_lower for c in CONTEXTO_EFECTIVA)

    if 'auna' in nombre_lower and len(limpiar_nombre(emisor)) < 12:
        return bool(re.search(r'\bauna\b', titulo_lower))

    if es_emisor_extranjero(emisor):
        variantes = variantes_emisor(emisor)
        termino = variantes[0].lower() if variantes else nombre_lower
        return termino in titulo_lower

    if es_emisor_soberano(emisor):
        return any(t in titulo_lower for t in TEMAS_FINANCIEROS_SOBERANOS)

    if 'banco de la nacion' in nombre_lower or 'banco de la nación' in nombre_lower:
        for pais in PAISES_EXTRANJEROS:
            if pais in titulo_lower:
                return False
        return True

    if 'rutas de lima' in nombre_lower:
        CONTEXTO_RUTAS = [
            'rutas de lima','concesion','concesión','peaje',
            'deuda','bonos','financiamiento','rating','calificacion',
            'moody','fitch','resultado','utilidad','inversion',
            'proyecto vial','deuda de largo plazo','brookfield','incumplimiento'
        ]
        return any(c in titulo_lower for c in CONTEXTO_RUTAS)

    nombres = variantes_emisor(emisor)
    termino_exacto = nombres[0].lower() if nombres else nombre_lower
    alias_tokens = [w.lower() for w in termino_exacto.split() if len(w) > 3][:2]
    if len(alias_tokens) >= 2:
        matches_alias = sum(1 for t in alias_tokens if t in titulo_lower)
        if matches_alias < len(alias_tokens):
            return False

    for pais in ['argentina', 'colombi', 'bogota', 'bogotá']:
        if pais in titulo_lower:
            if termino_exacto not in titulo_lower:
                return False

    return True

def variantes_emisor(emisor):
    emisor_clean = limpiar_nombre(emisor)
    emisor_clean = re.sub(r'\(.*?\)', '', emisor_clean).strip()
    sufijos = (r'\b(S\.A\.A\.|S\.A\.C\.|S\.A\.R\.L\.|S\.A\.|S\.A|SAA|SAC|S\.A\.B\.|S\.A\.B|'
               r'Corp\.?|Ltd\.?|Inc\.?|Co\.?|Perú|Peru|del Perú|de Peru|'
               r'Shopping Center|Centro Comercial|Administradora|Administrador|'
               r'Asset Management|Investments?|Fund|Luxembourg|Ireland|Europe|SICAV|LLP|GP)\b')
    limpio = re.sub(sufijos, '', emisor_clean, flags=re.IGNORECASE).strip().strip('.,&')
    limpio = re.sub(r'\s+', ' ', limpio).strip()

    variantes = []
    emisor_lower = emisor_clean.lower()

    mejor_alias, mejor_len = None, 0
    for clave, alias in ALIAS_EMISORES.items():
        if clave.strip() in emisor_lower and len(clave) > mejor_len:
            mejor_alias = alias
            mejor_len   = len(clave)
    if mejor_alias:
        variantes.append(mejor_alias)

    if limpio and limpio not in variantes:
        variantes.append(limpio)

    palabras = [p for p in limpio.split() if p.lower() not in PREFIJOS_IGNORAR]
    if len(palabras) >= 2:
        dos = f"{palabras[0]} {palabras[1]}"
        if dos not in variantes:
            variantes.append(dos)

    if palabras:
        primera = palabras[0]
        if primera not in variantes and len(primera) > 4 and primera.lower() not in TOKENS_GENERICOS:
            variantes.append(primera)

    return variantes if variantes else [emisor_clean]

def tokens_significativos(emisor):
    emisor_clean = limpiar_nombre(emisor)
    emisor_clean = re.sub(r'\(.*?\)', '', emisor_clean).strip()
    sufijos = (r'\b(S\.A\.A\.|S\.A\.C\.|S\.A\.|S\.A|SAA|SAC|S\.A\.B\.|'
               r'Corp\.?|Ltd\.?|Inc\.?|Co\.?|Perú|Peru|'
               r'Shopping Center|Centro Comercial|Administradora|'
               r'Asset Management|Luxembourg|Ireland|Europe|SICAV|LLP|GP)\b')
    limpio = re.sub(sufijos, '', emisor_clean, flags=re.IGNORECASE)
    return [w.lower().strip('.,&') for w in limpio.split()
            if w.lower().strip('.,&') not in PALABRAS_IGNORAR
            and len(w.strip('.,&')) > 2
            and w.lower().strip('.,&') not in TOKENS_GENERICOS]

def calcular_minimo(tokens):
    if not tokens: return 1
    if len(tokens) == 1: return 1
    if tokens[0] in TOKENS_GENERICOS: return min(2, len(tokens))
    if len(tokens) <= 2: return 1
    return 2

def usa_alias(emisor):
    nombres = variantes_emisor(emisor)
    return bool(nombres) and nombres[0] in ALIAS_EMISORES.values()

# ============================================================
# 7. FUENTES PRIORITARIAS
# ============================================================
FUENTES_PRIORITARIAS = [
    {"nombre": "BVL Oficial", "url": "https://www.bvl.com.pe/emisores/noticias-emisores"},
    {"nombre": "SMV Diario",  "url": "https://www.smv.gob.pe/SIMV/frm_hechosdeImportanciaDia?data=38C2EC33FA106691BB5B5039DACFDF50795D8EC3AF"},
]

def cargar_fuentes_prioritarias():
    print("  Cargando fuentes prioritarias (BVL, SMV)...")
    for fuente in FUENTES_PRIORITARIAS:
        try:
            req = urllib.request.Request(fuente["url"], headers=HEADERS)
            with urllib.request.urlopen(req, context=contexto, timeout=10) as res:
                raw = res.read().decode('utf-8', errors='ignore')
            raw = re.sub(r'<script[\s\S]*?</script>', '', raw, flags=re.IGNORECASE)
            raw = re.sub(r'<style[\s\S]*?</style>',   '', raw, flags=re.IGNORECASE)
            links = re.findall(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>([\s\S]*?)</a>', raw, re.IGNORECASE)
            for href, texto in links:
                texto_limpio = re.sub(r'<[^>]+>', ' ', texto).strip()
                texto_limpio = re.sub(r'\s+', ' ', texto_limpio).strip()
                if len(texto_limpio) < 15 or len(texto_limpio) > 300: continue
                if not href.startswith('http'):
                    base = re.match(r'https?://[^/]+', fuente["url"])
                    href = (base.group(0) if base else "") + "/" + href.lstrip("/")
                cache_prioritarias.append({
                    "titulo": texto_limpio, "fuente": fuente["nombre"],
                    "link": href, "fecha": "Hoy", "alerta": es_alerta_real(texto_limpio)
                })
            for linea in raw.split('\n'):
                linea = re.sub(r'<[^>]+>', ' ', linea)
                linea = re.sub(r'\s+', ' ', linea).strip()
                if len(linea) < 20 or len(linea) > 300: continue
                cache_prioritarias.append({
                    "titulo": linea, "fuente": fuente["nombre"],
                    "link": fuente["url"], "fecha": "Hoy", "alerta": es_alerta_real(linea)
                })
            print(f"    ✓ {fuente['nombre']}: OK")
        except Exception as e:
            print(f"    ✗ {fuente['nombre']} no disponible: {e}")

# ============================================================
# 8. BUSCAR EN CACHE BVL/SMV
# ============================================================
def buscar_en_cache_prioritarias(emisor):
    if es_emisor_extranjero(emisor): return []
    tokens  = tokens_significativos(emisor)
    nombres = variantes_emisor(emisor)
    termino_principal = nombres[0].lower() if nombres else limpiar_nombre(emisor).lower()
    minimo  = calcular_minimo(tokens)
    alias   = usa_alias(emisor)
    resultados, vistos = [], set()
    for entrada in cache_prioritarias:
        titulo_lower = entrada["titulo"].lower()
        if alias:
            if termino_principal not in titulo_lower: continue
        else:
            matches = sum(1 for t in tokens if t in titulo_lower)
            if matches < minimo: continue
            if len(termino_principal) > 5 and termino_principal not in titulo_lower:
                if matches < 2: continue
        if not titulo_relevante_para_emisor(entrada["titulo"], emisor): continue
        key = entrada["titulo"][:60]
        if key not in vistos:
            vistos.add(key)
            resultados.append(entrada)
    return resultados

# ============================================================
# 9. BLOOMBERG LÍNEA
# ============================================================
def buscar_bloomberg_linea(emisor):
    resultados = []
    nombres = variantes_emisor(emisor)
    termino = f'"{nombres[0]}"' if nombres else f'"{limpiar_nombre(emisor)}"'
    tokens  = tokens_significativos(emisor)
    minimo  = calcular_minimo(tokens)
    alias   = usa_alias(emisor)
    try:
        q   = urllib.parse.quote(f'{termino} site:bloomberglinea.com')
        url = f"https://news.google.com/rss/search?q={q}&hl=es-419&gl=PE&ceid=PE:es-419&tbs=qdr:m,sbd:1"
        req = urllib.request.Request(url, headers=HEADERS)
        with urllib.request.urlopen(req, context=contexto, timeout=10) as res:
            data = res.read().decode('utf-8', errors='ignore')
        for item in re.findall(r'<item>([\s\S]*?)</item>', data)[:5]:
            t = re.search(r'<title><!\[CDATA\[(.*?)\]\]></title>|<title>(.*?)</title>', item)
            l = re.search(r'<link>(.*?)</link>', item)
            d = re.search(r'<pubDate>(.*?)</pubDate>', item)
            titulo = (t.group(1) or t.group(2) or "").strip() if t else ""
            link   = l.group(1).strip() if l else "#"
            fecha  = d.group(1).strip() if d else ""
            if not titulo or not es_reciente(fecha): continue
            if not alias:
                matches = sum(1 for tk in tokens if tk in titulo.lower())
                if matches < minimo: continue
            if not titulo_relevante_para_emisor(titulo, emisor): continue
            resultados.append({
                "titulo": titulo, "fuente": "Bloomberg Línea",
                "link": link, "fecha": formatear_fecha(fecha),
                "alerta": es_alerta_real(titulo)
            })
    except Exception:
        pass
    return resultados

# ============================================================
# 10. GOOGLE NEWS
# ============================================================
def buscar_google_news(emisor):
    resultados, vistos = [], set()
    nombres   = variantes_emisor(emisor)
    tokens    = tokens_significativos(emisor)
    minimo    = calcular_minimo(tokens)
    con_alias = usa_alias(emisor)
    termino_base   = nombres[0] if nombres else limpiar_nombre(emisor)
    termino_exacto = f'"{termino_base}"'

    queries = [
        (f'{termino_exacto} (downgrade OR upgrade OR outlook OR rating OR Moody OR Fitch OR "S&P" OR "Pacific Credit Rating" OR "Apoyo & Asociados" OR perspectiva OR calificacion OR "deuda de largo plazo" OR bonos OR sindicado OR multa OR fusion OR adquisicion OR default)', True),
        (f'{termino_exacto} (finanzas OR resultados OR BVL OR SMV OR inversión OR utilidad OR "notas de deuda" OR capex OR ebitda OR trimestre OR dividendo OR "hecho de importancia")', False),
        (termino_exacto, False),
    ]
    for query, es_credit in queries:
        if len(resultados) >= 3: break
        try:
            url = f"https://news.google.com/rss/search?q={urllib.parse.quote(query)}&hl=es-419&gl=PE&ceid=PE:es-419&tbs=qdr:m,sbd:1"
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, context=contexto, timeout=10) as res:
                data = res.read().decode('utf-8', errors='ignore')
            for item in re.findall(r'<item>([\s\S]*?)</item>', data)[:8]:
                t = re.search(r'<title><!\[CDATA\[(.*?)\]\]></title>|<title>(.*?)</title>', item)
                l = re.search(r'<link>(.*?)</link>', item)
                d = re.search(r'<pubDate>(.*?)</pubDate>', item)
                s = re.search(r'<source[^>]*>(.*?)</source>', item)
                titulo = (t.group(1) or t.group(2) or "").strip() if t else ""
                link   = l.group(1).strip() if l else "#"
                fecha  = d.group(1).strip() if d else ""
                fuente = (s.group(1) or "Prensa").strip() if s else "Prensa"
                if not titulo or titulo in vistos: continue
                if not es_reciente(fecha): continue
                if not con_alias:
                    matches = sum(1 for tk in tokens if tk in titulo.lower())
                    if matches < minimo: continue
                if not titulo_relevante_para_emisor(titulo, emisor): continue
                vistos.add(titulo)
                resultados.append({
                    "titulo": titulo, "fuente": fuente,
                    "link": link, "fecha": formatear_fecha(fecha),
                    "alerta": es_alerta_real(titulo)
                })
        except Exception:
            pass
    return resultados[:3]

# ============================================================
# 11. RESUMEN IA POR EMISOR
# ============================================================
PROMPT_CRITERIOS = """Criterios de priorización (en orden):
1. Downgrade/default/estrés financiero/restricción rescates fondos
2. Fusiones/adquisiciones/cambios de control
3. Upgrade/outlook negativo/resultados fuera expectativas/multas graves
4. Outlook positivo/informes clasificadoras/nuevas emisiones deuda
5. Cambios regulatorios/AUM fondos/perspectiva general
6. CAPEX/resultados financieros/cambios operativos fondos
7. Financiamiento/estructura accionaria
8. Analistas/expansión/proyectos generales"""

def generar_resumen_ia(emisor, noticias):
    if not noticias:
        return None
    titulos = "\n".join([f"- {n['titulo']} ({n['fecha']})" for n in noticias])
    prompt = f"""Eres un analista de inversiones senior. Basándote SOLO en estos titulares del emisor "{limpiar_nombre(emisor)}", genera un resumen ejecutivo en español de máximo 2 oraciones (60 palabras).

{PROMPT_CRITERIOS}

NO inventes. USA SOLO lo que está en los titulares.
Titulares:\n{titulos}\nResponde SOLO con el resumen."""
    try:
        payload = json.dumps({
            "model": "claude-sonnet-4-20250514",
            "max_tokens": 150,
            "messages": [{"role": "user", "content": prompt}]
        }).encode('utf-8')
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={"Content-Type": "application/json", "anthropic-version": "2023-06-01"},
            method="POST"
        )
        with urllib.request.urlopen(req, context=contexto, timeout=15) as res:
            data = json.loads(res.read().decode('utf-8'))
            return data.get("content", [{}])[0].get("text", "").strip()
    except Exception as e:
        print(f"     ⚠ Resumen IA: {e}")
        return None

# ============================================================
# 12. RESUMEN CONSOLIDADO (máximo 10, score >= 5)
# ============================================================
def generar_resumen_portafolio(resultados_por_segmento):
    noticias_criticas = []
    for seg, emisores in resultados_por_segmento.items():
        for emisor, noticias, resumen in emisores:
            for n in noticias:
                score = n.get("score", 0)
                alerta = n.get("alerta", False)
                if alerta or score >= SCORE_MINIMO_RESUMEN:
                    noticias_criticas.append({
                        "emisor":   limpiar_nombre(emisor),
                        "segmento": seg,
                        "titulo":   n["titulo"],
                        "link":     n["link"],
                        "fecha":    n["fecha"],
                        "fuente":   n["fuente"],
                        "score":    score,
                        "alerta":   alerta,
                    })

    noticias_criticas.sort(key=lambda x: -(10 if x["alerta"] else x["score"]))
    top = noticias_criticas[:10]

    if not top:
        return None, []

    lineas = "\n".join([
        f"- [{x['emisor']} / {x['segmento']}] {x['titulo']} ({x['fecha']})"
        for x in top
    ])

    prompt = f"""Eres un analista de inversiones senior de un fondo peruano.
Basándote SOLO en estos titulares del portafolio, genera un resumen ejecutivo en español de máximo 5 oraciones (150 palabras).

{PROMPT_CRITERIOS}

Menciona los emisores por nombre. NO inventes. USA SOLO los titulares.
Titulares:\n{lineas}\nResponde SOLO con el resumen ejecutivo."""

    resumen_texto = None
    try:
        payload = json.dumps({
            "model": "claude-sonnet-4-20250514",
            "max_tokens": 250,
            "messages": [{"role": "user", "content": prompt}]
        }).encode('utf-8')
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={"Content-Type": "application/json", "anthropic-version": "2023-06-01"},
            method="POST"
        )
        with urllib.request.urlopen(req, context=contexto, timeout=20) as res:
            data = json.loads(res.read().decode('utf-8'))
            resumen_texto = data.get("content", [{}])[0].get("text", "").strip()
    except Exception as e:
        print(f"  ⚠ Resumen portafolio: {e}")

    return resumen_texto, top

# ============================================================
# 13. CONSOLIDAR NOTICIAS (máximo 3)
# ============================================================
def obtener_noticias(emisor):
    todas = (
        buscar_en_cache_prioritarias(emisor) +
        buscar_bloomberg_linea(emisor) +
        buscar_google_news(emisor)
    )
    vistos, resultado = set(), []
    for n in todas:
        key = n["titulo"][:60].lower().strip()
        if key not in vistos:
            vistos.add(key)
            n["score"] = calcular_score(n["titulo"])
            resultado.append(n)
        if len(resultado) >= 3: break

    resultado.sort(key=lambda x: -(10 if x.get("alerta") else x.get("score", 0)))

    if len(resultado) < 3:
        nombres = variantes_emisor(emisor)
        termino = f'"{nombres[0]}"' if nombres else f'"{limpiar_nombre(emisor)}"'
        titulos_vistos = {n["titulo"] for n in resultado}
        for fq in [f'{termino} finanzas deuda bonos resultados Peru clasificacion',
                   f'{termino} financiero crediticio calificacion ebitda']:
            if len(resultado) >= 3: break
            try:
                url = f"https://news.google.com/rss/search?q={urllib.parse.quote(fq)}&hl=es-419&gl=PE&ceid=PE:es-419&tbs=qdr:m,sbd:1"
                req = urllib.request.Request(url, headers=HEADERS)
                with urllib.request.urlopen(req, context=contexto, timeout=10) as res:
                    data = res.read().decode('utf-8', errors='ignore')
                for item in re.findall(r'<item>([\s\S]*?)</item>', data)[:10]:
                    if len(resultado) >= 3: break
                    t = re.search(r'<title><!\[CDATA\[(.*?)\]\]></title>|<title>(.*?)</title>', item)
                    l = re.search(r'<link>(.*?)</link>', item)
                    d = re.search(r'<pubDate>(.*?)</pubDate>', item)
                    s = re.search(r'<source[^>]*>(.*?)</source>', item)
                    titulo = (t.group(1) or t.group(2) or "").strip() if t else ""
                    link   = l.group(1).strip() if l else "#"
                    fecha  = d.group(1).strip() if d else ""
                    fuente = (s.group(1) or "Prensa").strip() if s else "Prensa"
                    if not titulo or titulo in titulos_vistos: continue
                    if not titulo_es_relevante_financiero(titulo): continue
                    if not titulo_relevante_para_emisor(titulo, emisor): continue
                    if not es_reciente(fecha): continue
                    titulo_lower = titulo.lower()
                    tiene_fin = any(kw in titulo_lower for kw in [
                        'financier','econom','banco','credito','crédito','inversion',
                        'inversión','deuda','bonos','resultado','utilidad','ganancia',
                        'ebitda','trimestre','rating','calificacion','moody','fitch',
                        'bolsa','accion','deuda de largo plazo','notas de deuda',
                        'fusion','adquisicion','dividendo','capex'])
                    if not tiene_fin: continue
                    titulos_vistos.add(titulo)
                    resultado.append({
                        "titulo": titulo, "fuente": fuente,
                        "link": link, "fecha": formatear_fecha(fecha),
                        "alerta": es_alerta_real(titulo), "score": calcular_score(titulo)
                    })
            except Exception:
                pass
        resultado.sort(key=lambda x: -(10 if x.get("alerta") else x.get("score", 0)))

    return resultado[:3]

# ============================================================
# 14. HELPERS HTML
# ============================================================
ORDEN_SEG = ["Renta Fija", "Renta Variable", "Fondos", "Alternativos"]
COLOR_SEG = {
    "Renta Fija":     ("text-blue-400",    "border-blue-500/30",    "bg-blue-500/10"),
    "Renta Variable": ("text-purple-400",  "border-purple-500/30",  "bg-purple-500/10"),
    "Fondos":         ("text-emerald-400", "border-emerald-500/30", "bg-emerald-500/10"),
    "Alternativos":   ("text-amber-400",   "border-amber-500/30",   "bg-amber-500/10"),
}
TAB_META = {
    "Renta Fija":     ("💼", "text-blue-300",    "border-blue-500/50",    "bg-blue-500/20"),
    "Renta Variable": ("📈", "text-purple-300",  "border-purple-500/50",  "bg-purple-500/20"),
    "Fondos":         ("🏦", "text-emerald-300", "border-emerald-500/50", "bg-emerald-500/20"),
    "Alternativos":   ("🔷", "text-amber-300",   "border-amber-500/50",   "bg-amber-500/20"),
}

def badge_fuente(fuente):
    f = fuente.lower()
    if "bvl"       in f: return "bg-blue-900/40 text-blue-300 border-blue-600/40"
    if "smv"       in f: return "bg-indigo-900/40 text-indigo-300 border-indigo-600/40"
    if "bloomberg" in f: return "bg-orange-900/40 text-orange-300 border-orange-600/40"
    return "bg-gray-800 text-gray-400 border-gray-700"

def dot_color(n):
    score = n.get("score", 0)
    if n.get("alerta") or score >= 10: return "bg-red-500"
    if score >= 9:                      return "bg-purple-500"
    if score >= 8:                      return "bg-orange-500"
    if score >= 7:                      return "bg-blue-400"
    if score >= 5:                      return "bg-yellow-500"
    if "bvl"       in n["fuente"].lower(): return "bg-blue-500"
    if "smv"       in n["fuente"].lower(): return "bg-indigo-500"
    if "bloomberg" in n["fuente"].lower(): return "bg-orange-400"
    return "bg-gray-500"

def render_cards(emisores_con_noticias, seg):
    tc, bc, bgc = COLOR_SEG.get(seg, ("text-gray-400", "border-gray-600", "bg-gray-800"))
    cards = []
    for emisor, noticias, resumen in emisores_con_noticias:
        nombre_display = limpiar_nombre(emisor)
        tiene_alerta   = any(n.get("alerta") for n in noticias)
        color_riesgo   = get_semaforo(emisor)
        badge_riesgo   = semaforo_html(color_riesgo)
        es_prior       = es_emisor_prioritario(emisor)

        if tiene_alerta:
            cb = "border-red-500/50 shadow-red-950/30 shadow-lg"
            bg = "bg-gradient-to-b from-gray-900 to-red-950/15"
        elif color_riesgo == "rojo":
            cb = "border-red-500/30 shadow-md"
            bg = "bg-gradient-to-b from-gray-900 to-red-950/10"
        elif color_riesgo == "amarillo":
            cb = "border-yellow-500/30 shadow-md"
            bg = "bg-gradient-to-b from-gray-900 to-yellow-950/10"
        else:
            cb = "border-gray-800"
            bg = "bg-gray-900"

        pin_badge    = '<span class="text-[9px] bg-indigo-500/20 border border-indigo-500/30 text-indigo-300 px-1.5 py-0.5 rounded font-bold">📌 Prioritario</span>' if es_prior else ""
        rating_badge = '<span class="text-[9px] bg-red-500/20 border border-red-500/30 text-red-400 px-2 py-0.5 rounded font-bold animate-pulse">⚠ Rating Alert</span>' if tiene_alerta else ""

        resumen_html = ""
        if resumen:
            resumen_html = f"""
            <div class="bg-gray-800/60 border border-gray-700/50 rounded-lg p-2.5 mb-1">
              <span class="text-[9px] bg-violet-500/20 border border-violet-500/30 text-violet-300 px-1.5 py-0.5 rounded font-bold">✦ Resumen IA</span>
              <p class="text-[11px] text-gray-300 leading-snug italic mt-1">{html.escape(resumen)}</p>
            </div>"""

        noticias_html = []
        for n in noticias:
            dc  = dot_color(n)
            ibg = "bg-red-950/30 border-red-500/40" if n.get("alerta") else "bg-gray-800/40 border-gray-700/40"
            itx = "text-red-200" if n.get("alerta") else "text-gray-300"
            fb  = badge_fuente(n["fuente"])
            noticias_html.append(f"""
              <div class="border rounded-lg p-2.5 {ibg}">
                <div class="flex gap-2 items-start">
                  <span class="mt-1.5 w-2 h-2 rounded-full shrink-0 {dc}"></span>
                  <div class="flex-1 min-w-0">
                    <a href="{html.escape(n['link'])}" target="_blank"
                       class="text-xs font-medium {itx} hover:text-blue-400 leading-snug line-clamp-2">
                      {html.escape(n['titulo'])}
                    </a>
                    <div class="flex justify-between items-center mt-1.5 gap-2">
                      <span class="text-[10px] border px-1.5 py-0.5 rounded font-mono {fb}">{html.escape(n['fuente'])}</span>
                      <span class="text-[10px] text-gray-500 font-mono shrink-0">{html.escape(n['fecha'])}</span>
                    </div>
                  </div>
                </div>
              </div>""")

        cards.append(f"""
          <div class="rounded-xl border p-5 flex flex-col gap-3 {cb} {bg}">
            <div class="flex justify-between items-start border-b border-gray-800/70 pb-2.5 gap-2">
              <h3 class="text-sm font-bold uppercase tracking-wide leading-tight">{html.escape(nombre_display)}</h3>
              <div class="flex flex-col items-end gap-1 shrink-0">
                <span class="text-[10px] border px-2 py-0.5 rounded {tc} {bc} {bgc}">{html.escape(seg)}</span>
                {badge_riesgo}{pin_badge}{rating_badge}
              </div>
            </div>
            {resumen_html}
            <div class="flex flex-col gap-2">{''.join(noticias_html)}</div>
          </div>""")
    return ''.join(cards)

def render_tab_todos(resumen_portafolio, noticias_criticas_top):
    if not noticias_criticas_top and not resumen_portafolio:
        return '<div class="text-gray-500 text-sm italic p-8 text-center">Sin noticias materiales estas 2 semanas.</div>'

    resumen_bloque = ""
    if resumen_portafolio:
        resumen_bloque = f"""
      <div class="mb-5 pb-5 border-b border-gray-700/50">
        <div class="flex items-center gap-2 mb-2">
          <span class="text-[10px] bg-violet-500/20 border border-violet-500/30 text-violet-300 px-2 py-0.5 rounded font-bold">✦ Análisis IA del Portafolio</span>
        </div>
        <p class="text-sm text-gray-200 leading-relaxed">{html.escape(resumen_portafolio)}</p>
      </div>"""

    items_html = ""
    for n in noticias_criticas_top:
        score = n.get("score", 0)
        alerta = n.get("alerta", False)
        icono, tipo = SCORE_LABELS.get(score, ("•", "General"))
        if alerta or score >= 10:
            row_bg, txt = "bg-red-950/20 border-red-500/30", "text-red-200"
        elif score >= 9:
            row_bg, txt = "bg-purple-950/20 border-purple-500/30", "text-purple-200"
        elif score >= 8:
            row_bg, txt = "bg-orange-950/20 border-orange-500/30", "text-orange-200"
        elif score >= 7:
            row_bg, txt = "bg-blue-950/20 border-blue-500/30", "text-blue-200"
        elif score >= 5:
            row_bg, txt = "bg-yellow-950/10 border-yellow-600/20", "text-yellow-100"
        else:
            row_bg, txt = "bg-gray-800/30 border-gray-700/30", "text-gray-300"

        items_html += f"""
        <div class="flex gap-3 items-start p-3 rounded-lg border {row_bg}">
          <span class="text-base shrink-0 mt-0.5">{icono}</span>
          <div class="flex-1 min-w-0">
            <div class="flex items-center gap-2 mb-0.5 flex-wrap">
              <span class="text-[10px] font-bold text-gray-300 uppercase tracking-wide">{html.escape(n['emisor'])}</span>
              <span class="text-[9px] text-gray-600">·</span>
              <span class="text-[9px] text-gray-500">{html.escape(n['segmento'])}</span>
              <span class="text-[9px] text-gray-600">·</span>
              <span class="text-[9px] bg-gray-800 text-gray-400 px-1.5 py-0.5 rounded border border-gray-700">{tipo}</span>
            </div>
            <a href="{html.escape(n['link'])}" target="_blank"
               class="text-xs font-medium {txt} hover:text-blue-400 leading-snug line-clamp-2">
              {html.escape(n['titulo'])}
            </a>
            <div class="flex gap-2 mt-1 text-[10px] text-gray-600">
              <span>{html.escape(n['fuente'])}</span>
              <span>·</span>
              <span>{html.escape(n['fecha'])}</span>
            </div>
          </div>
        </div>"""

    return f"""
  <div class="max-w-4xl mx-auto">
    <div class="bg-gray-900/80 border border-gray-700/50 rounded-xl p-6 mb-8">
      <h2 class="text-base font-bold text-white mb-4 flex items-center gap-2">
        <span class="w-2 h-2 rounded-full bg-violet-400 inline-block"></span>
        Resumen Ejecutivo del Portafolio
        <span class="text-[10px] text-gray-500 font-normal ml-1">— Últimas 2 semanas · Top {len(noticias_criticas_top)} eventos materiales</span>
      </h2>
      {resumen_bloque}
      <h3 class="text-xs font-bold text-gray-400 uppercase tracking-wider mb-3">
        Noticias Materiales
      </h3>
      <div class="flex flex-col gap-2">{items_html}</div>
    </div>
  </div>"""

# ============================================================
# 15. EJECUCIÓN PRINCIPAL
# ============================================================
cargar_fuentes_prioritarias()
emisores_data = leer_emisores_excel("Emisores.xlsx")

if not emisores_data:
    print("ADVERTENCIA: No se leyeron emisores. Verifica Emisores.xlsx")
else:
    print(f"\n📋 {len(emisores_data)} emisores cargados desde Excel:")
    for e, s in emisores_data.items():
        v      = variantes_emisor(e)
        color  = get_semaforo(e)
        sem    = f"[{color.upper()}]" if color else "[sin semáforo]"
        prior  = "⭐" if es_emisor_prioritario(e) else ""
        extran = "🌍" if es_emisor_extranjero(e) else ""
        sobera = "🏛" if es_emisor_soberano(e) else ""
        print(f"   {limpiar_nombre(e):50s} → {v[0]:30s} {sem} {prior}{extran}{sobera}")

segmentos = {}
for emisor, seg in emisores_data.items():
    seg_norm = seg.strip()
    segmentos.setdefault(seg_norm, []).append(emisor)

segs_ordenados = sorted(segmentos.keys(), key=lambda s: ORDEN_SEG.index(s) if s in ORDEN_SEG else 99)
print(f"\nSegmentos: {list(segmentos.keys())}")

# ============================================================
# 16. PRE-CALCULAR NOTICIAS + RESÚMENES IA
# ============================================================
print("\n🔍 Buscando noticias y generando resúmenes...")
resultados_por_segmento = {}

for seg in segs_ordenados:
    emisores_con_noticias = []
    for emisor in segmentos[seg]:
        nombre_display = limpiar_nombre(emisor)
        v = variantes_emisor(emisor)
        print(f"  ▶ [{seg}] {nombre_display:45s} → {v[0]}")
        noticias = obtener_noticias(emisor)
        if noticias:
            resumen = generar_resumen_ia(emisor, noticias)
            emisores_con_noticias.append((emisor, noticias, resumen))
            print(f"     ✓ {len(noticias)} noticias | IA: {'✓' if resumen else '✗'}")
        else:
            print(f"     ↳ sin noticias, omitido")

    def sort_key_emisor(item):
        emisor, noticias, _ = item
        tiene_alerta = any(n.get("alerta") for n in noticias)
        mejor_score  = max((n.get("score", 0) for n in noticias), default=0)
        return (0 if es_emisor_prioritario(emisor) else 1,
                0 if tiene_alerta else 1, -mejor_score)

    emisores_con_noticias.sort(key=sort_key_emisor)
    if emisores_con_noticias:
        resultados_por_segmento[seg] = emisores_con_noticias
    else:
        print(f"  ⚠ Segmento '{seg}' sin noticias, omitido")

print("\n✦ Generando resumen ejecutivo del portafolio...")
resumen_portafolio, noticias_criticas_top = generar_resumen_portafolio(resultados_por_segmento)

# ============================================================
# 17. CONSTRUCCIÓN HTML CON TABS
# ============================================================
total_eventos = len(noticias_criticas_top)

tabs_nav = f"""
    <button onclick="mostrarTab('todos')" id="tab-btn-todos"
      class="tab-btn px-4 py-2 rounded-lg text-sm font-semibold border transition-all
             bg-white/10 border-white/20 text-white">
      📋 Resumen <span class="ml-1 text-[10px] opacity-60">({total_eventos} eventos)</span>
    </button>"""

for seg, items in resultados_por_segmento.items():
    icono, tc, bc, bgc = TAB_META.get(seg, ("•","text-gray-300","border-gray-600","bg-gray-800"))
    sid = seg.lower().replace(' ', '-')
    tabs_nav += f"""
    <button onclick="mostrarTab('{sid}')" id="tab-btn-{sid}"
      class="tab-btn px-4 py-2 rounded-lg text-sm font-semibold border transition-all
             {bgc} {bc} {tc} opacity-60 hover:opacity-100">
      {icono} {html.escape(seg)} <span class="ml-1 text-[10px] opacity-60">({len(items)})</span>
    </button>"""

tabs_content = f'<div id="tab-todos" class="tab-panel">{render_tab_todos(resumen_portafolio, noticias_criticas_top)}</div>'

for seg, items in resultados_por_segmento.items():
    tc  = COLOR_SEG.get(seg, ("text-gray-400",))[0]
    sid = seg.lower().replace(' ', '-')
    tabs_content += f"""
<div id="tab-{sid}" class="tab-panel hidden">
  <h2 class="text-lg font-bold mb-5 pb-2 border-b border-gray-800 {tc}">{html.escape(seg)}</h2>
  <div class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-5">{render_cards(items, seg)}</div>
</div>"""

page = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Monitor Crediticio — Portafolio</title>
  <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-950 text-gray-100 min-h-screen font-sans">
<div class="max-w-7xl mx-auto px-4 py-10">

  <header class="mb-10 border-b border-gray-800 pb-6 flex flex-col md:flex-row md:justify-between md:items-end gap-4">
    <div>
      <h1 class="text-3xl font-extrabold bg-gradient-to-r from-blue-400 via-indigo-400 to-emerald-400 bg-clip-text text-transparent">
        Monitor Crediticio del Portafolio
      </h1>
      <p class="text-gray-400 text-sm mt-1">
        Fuentes: BVL · SMV · Bloomberg Línea · Google News &nbsp;·&nbsp; Últimas 2 semanas
      </p>
    </div>
    <div class="text-xs text-gray-500 font-mono">Actualizado: {fecha_reporte}</div>
  </header>

  <div class="flex flex-wrap gap-3 mb-6 text-xs">
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2 h-2 rounded-full bg-red-500"></span> Downgrade/Default/Estrés
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2 h-2 rounded-full bg-purple-500"></span> Fusión/Adquisición
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2 h-2 rounded-full bg-orange-500"></span> Upgrade/Outlook Neg./Multa
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2 h-2 rounded-full bg-blue-400"></span> Clasificadora/Emisión Deuda
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2 h-2 rounded-full bg-yellow-500"></span> Perspectiva/Regulatorio
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2 h-2 rounded-full bg-gray-500"></span> General
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      🔴 Riesgo Alto &nbsp;·&nbsp; 🟡 Moderado &nbsp;·&nbsp; 🟢 Bajo
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="text-violet-300">✦</span> Resumen IA
    </span>
  </div>

  <div class="flex flex-wrap gap-2 mb-8 border-b border-gray-800 pb-4">
    {tabs_nav}
  </div>

  {tabs_content}

  <footer class="mt-12 pt-6 border-t border-gray-800 text-center text-xs text-gray-600">
    Monitor Crediticio · BVL · SMV · Bloomberg Línea · Google News · Solo fuentes públicas
  </footer>
</div>

<script>
  function mostrarTab(id) {{
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.add('hidden'));
    document.querySelectorAll('.tab-btn').forEach(b => {{
      b.classList.remove('opacity-100');
      b.classList.add('opacity-60');
    }});
    const panel = document.getElementById('tab-' + id);
    if (panel) panel.classList.remove('hidden');
    const btn = document.getElementById('tab-btn-' + id);
    if (btn) {{
      btn.classList.remove('opacity-60');
      btn.classList.add('opacity-100');
    }}
  }}
  mostrarTab('todos');
</script>
</body>
</html>"""

# ============================================================
# 18. GUARDAR
# ============================================================
with open("index.html", "w", encoding="utf-8") as f:
    f.write(page)

print(f"\n✅ index.html generado — {len(emisores_data)} emisores procesados.")
