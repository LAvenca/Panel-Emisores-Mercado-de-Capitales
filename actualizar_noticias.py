import urllib.request
import urllib.parse
import re
import ssl
import html
import html as html_mod
import zipfile
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
# 2. SEMÁFORO DE RIESGO — edita aquí manualmente
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
    "SCOTIABANK PERU S.A.A.":              "verde",
    "MIBANCO":                              "verde",
    "BANCO GNB PERU S.A.":                  "amarillo",
    "FINANCIERA EFECTIVA S.A.":             "amarillo",
    "CERRO VERDE":                          "verde",
    "SOUTHERN PERU COPPER CORPORATION":     "verde",
    "ACEROS AREQUIPA":                      "verde",
    "INRETAIL PERU CORP.":                  "verde",
    "ADMINISTRADORA JOCKEY PLAZA S.A.":     "amarillo",
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
# 3. PRIORIDAD DE EMISORES Y SCORING DE NOTICIAS
# ============================================================
EMISORES_PRIORITARIOS = [
    "rutas de lima",
    "auna",
]

SCORING_NOTICIAS = [
    (5, [
        'downgrade', 'upgrade', 'rebaja', 'mejora', 'calificacion', 'calificación',
        'clasificacion', 'clasificación', 'rating', 'moody', 'fitch', 's&p',
        'standard & poor', 'investment grade', 'grado de inversion', 'grado de inversión',
        'watch negative', 'watch positive', 'creditwatch', 'outlook', 'perspectiva',
        'bajo revision', 'bajo revisión', 'cambio de perspectiva', 'alza de nota',
        'baja de nota', 'afirmacion', 'afirmación',
    ]),
    (4, [
        'multa', 'sancion', 'sanción', 'demanda', 'denuncia', 'investigacion',
        'investigación', 'fraude', 'corrupcion', 'corrupción', 'escandalo',
        'escándalo', 'indecopi', 'sbs', 'regulador', 'incumplimiento',
        'default', 'quiebra', 'concurso de acreedores', 'riesgo reputacional',
        'contingencia legal', 'proceso judicial',
    ]),
    (3, [
        'bonos', 'deuda', 'emision', 'emisión', 'sindicado', 'prestamo', 'préstamo',
        'credito', 'crédito', 'financiamiento', 'refinanciamiento', 'linea de credito',
        'línea de crédito', 'aumento de capital', 'suscripcion', 'suscripción',
        'colocacion', 'colocación', 'oferta publica', 'oferta pública', 'spread',
        'tasa de interes', 'tasa de interés', 'cupón', 'cupon', 'vencimiento',
    ]),
    (2, [
        'capex', 'inversion', 'inversión', 'proyecto', 'expansion', 'expansión',
        'planta', 'infraestructura', 'contrato', 'concesion', 'concesión',
        'adjudicacion', 'adjudicación', 'licitacion', 'licitación', 'obra',
        'ampliacion', 'ampliación', 'construccion', 'construcción',
    ]),
    (1, [
        'utilidad', 'utilidades', 'ganancia', 'perdida', 'pérdida', 'resultado',
        'resultados', 'ebitda', 'ingresos', 'revenue', 'trimestre', 'semestre',
        'anual', 'balance', 'flujo de caja', 'cash flow', 'margen', 'rentabilidad',
        'dividendo', 'dividendos', 'acciones', 'bolsa', 'bvl',
    ]),
]

def calcular_score(titulo):
    titulo_lower = titulo.lower()
    for score, keywords in SCORING_NOTICIAS:
        if any(k in titulo_lower for k in keywords):
            return score
    return 0

def es_emisor_prioritario(emisor):
    nombre_lower = limpiar_nombre(emisor).lower()
    return any(p in nombre_lower for p in EMISORES_PRIORITARIOS)

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
    'downgrade', 'upgrade', 'outlook', 'perspectiva', 'rating', 'calificacion',
    'calificación', 'moody', 'fitch', 's&p', 'standard & poor', 'investment grade',
    'grado de inversion', 'grado de inversión', 'watch negative', 'watch positive',
    'creditwatch', 'bajo revision', 'bajo revisión', 'rebaja', 'mejora crediticia',
    'deuda', 'bonos', 'emision', 'emisión', 'sindicado', 'hecho de importancia',
    'aumento de capital', 'soberano', 'riesgo pais', 'riesgo país', 'spread'
]

PALABRAS_IGNORAR = {
    'de', 'la', 'el', 'los', 'las', 'del', 'y', 'en', 'al',
    'sa', 'saa', 'sac', 'sab', 'the', 'of', 'and'
}

TOKENS_GENERICOS = {
    'banco', 'fondo', 'corp', 'group', 'grupo', 'financiera',
    'inversiones', 'holding', 'capital', 'asset', 'trust'
}

PREFIJOS_IGNORAR = {
    'administradora', 'administrador', 'compania', 'compañia',
    'empresa', 'grupo', 'corporacion', 'corporación', 'sociedad'
}

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
    "southern peru":       "Southern Peru",
    "volcan":              "Volcan",
    "alicorp":             "Alicorp",
    "aceros arequipa":     "Aceros Arequipa",
    "banco de la nacion":  "Banco de la Nación",
    "banco gnb":           "Banco GNB",
    "mibanco":             "Mibanco",
    "financiera efectiva": "Financiera Efectiva",
    "bbva":                "BBVA Peru",
    "scotiabank":          "Scotiabank Peru",
    "bcp":                 "BCP",
    "interbank":           "Interbank",
    "brown brothers":      "Brown Brothers Harriman",
    "fibra prime":         "Fibra Prime",
    "fossal":              "Fossal",
    "laive":               "Laive",
    "cementos pacasmayo":  "Cementos Pacasmayo",
    "ferreyros":           "Ferreyros",
    "luz del sur":         "Luz del Sur",
    "pluz energia":        "Pluz Energía Peru",
    "endispc":             "Enel Distribución Peru",
    "enel distribuc":      "Enel Distribución Peru",
    "enel":                "Enel Peru",
    "edelnor":             "Enel Peru",
    "telefonica":          "Telefónica del Perú",
    "entel":               "Entel Peru",
    "lima airport":        "LAP",
    "lap":                 "LAP",
    "rutas de lima":       "Rutas de Lima",
    "auna":                "Auna",
}

LIMITE_FECHA  = datetime.now() - timedelta(days=7)
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
    meses = {
        "Jan":"Ene","Feb":"Feb","Mar":"Mar","Apr":"Abr","May":"May",
        "Jun":"Jun","Jul":"Jul","Aug":"Ago","Sep":"Sep","Oct":"Oct",
        "Nov":"Nov","Dec":"Dic"
    }
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
# 6. UTILIDADES DE NOMBRE
# ============================================================
def limpiar_nombre(emisor):
    nombre = html_mod.unescape(emisor)
    return re.sub(r'\s+', ' ', nombre).strip()

def variantes_emisor(emisor):
    emisor_clean = limpiar_nombre(emisor)
    emisor_clean = re.sub(r'\(.*?\)', '', emisor_clean).strip()
    sufijos = (
        r'\b(S\.A\.A\.|S\.A\.C\.|S\.A\.|S\.A|SAA|SAC|S\.A\.B\.|S\.A\.B|'
        r'Corp\.?|Ltd\.?|Inc\.?|Co\.?|Perú|Peru|del Perú|de Peru|'
        r'Shopping Center|Centro Comercial|Administradora|Administrador)\b'
    )
    limpio = re.sub(sufijos, '', emisor_clean, flags=re.IGNORECASE).strip().strip('.,&')
    limpio = re.sub(r'\s+', ' ', limpio).strip()
    variantes = []
    emisor_lower = emisor_clean.lower()
    for clave, alias in ALIAS_EMISORES.items():
        if clave.strip() in emisor_lower:
            variantes.append(alias)
            break
    if limpio and limpio not in variantes:
        variantes.append(limpio)
    palabras = [p for p in limpio.split() if p.lower() not in PREFIJOS_IGNORAR]
    if len(palabras) >= 2:
        dos = f"{palabras[0]} {palabras[1]}"
        if dos not in variantes:
            variantes.append(dos)
    if palabras:
        primera = palabras[0]
        if primera not in variantes and len(primera) > 3:
            variantes.append(primera)
    return variantes

def tokens_significativos(emisor):
    emisor_clean = limpiar_nombre(emisor)
    emisor_clean = re.sub(r'\(.*?\)', '', emisor_clean).strip()
    sufijos = (
        r'\b(S\.A\.A\.|S\.A\.C\.|S\.A\.|S\.A|SAA|SAC|S\.A\.B\.|'
        r'Corp\.?|Ltd\.?|Inc\.?|Co\.?|Perú|Peru|'
        r'Shopping Center|Centro Comercial|Administradora)\b'
    )
    limpio = re.sub(sufijos, '', emisor_clean, flags=re.IGNORECASE)
    return [
        w.lower().strip('.,&')
        for w in limpio.split()
        if w.lower().strip('.,&') not in PALABRAS_IGNORAR
        and len(w.strip('.,&')) > 2
    ]

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
# 7. FUENTES PRIORITARIAS: BVL Y SMV
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
                    "link": href, "fecha": "Hoy",
                    "alerta": bool(re.search('|'.join(PALABRAS_CREDITICIAS), texto_limpio, re.IGNORECASE))
                })
            for linea in raw.split('\n'):
                linea = re.sub(r'<[^>]+>', ' ', linea)
                linea = re.sub(r'\s+', ' ', linea).strip()
                if len(linea) < 20 or len(linea) > 300: continue
                cache_prioritarias.append({
                    "titulo": linea, "fuente": fuente["nombre"],
                    "link": fuente["url"], "fecha": "Hoy",
                    "alerta": bool(re.search('|'.join(PALABRAS_CREDITICIAS), linea, re.IGNORECASE))
                })
            print(f"    ✓ {fuente['nombre']}: OK")
        except Exception as e:
            print(f"    ✗ {fuente['nombre']} no disponible: {e}")

# ============================================================
# 8. BUSCAR EN CACHE BVL/SMV
# ============================================================
def buscar_en_cache_prioritarias(emisor):
    tokens  = tokens_significativos(emisor)
    nombres = variantes_emisor(emisor)
    termino_principal = nombres[0].lower() if nombres else limpiar_nombre(emisor).lower()
    minimo  = calcular_minimo(tokens)
    alias   = usa_alias(emisor)
    resultados = []
    vistos = set()
    for entrada in cache_prioritarias:
        titulo_lower = entrada["titulo"].lower()
        if alias:
            if termino_principal not in titulo_lower: continue
        else:
            matches = sum(1 for t in tokens if t in titulo_lower)
            if matches < minimo: continue
            if len(termino_principal) > 5 and termino_principal not in titulo_lower:
                if matches < 2: continue
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
    termino = nombres[0] if nombres else limpiar_nombre(emisor)
    tokens  = tokens_significativos(emisor)
    minimo  = calcular_minimo(tokens)
    alias   = usa_alias(emisor)
    try:
        q   = urllib.parse.quote(f'{termino} site:bloomberglinea.com')
        url = f"https://news.google.com/rss/search?q={q}&hl=es-419&gl=PE&ceid=PE:es-419&tbs=qdr:w,sbd:1"
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
            resultados.append({
                "titulo": titulo, "fuente": "Bloomberg Línea",
                "link": link, "fecha": formatear_fecha(fecha),
                "alerta": bool(re.search('|'.join(PALABRAS_CREDITICIAS), titulo, re.IGNORECASE))
            })
    except Exception:
        pass
    return resultados

# ============================================================
# 10. GOOGLE NEWS
# ============================================================
def buscar_google_news(emisor):
    resultados = []
    vistos    = set()
    nombres   = variantes_emisor(emisor)
    tokens    = tokens_significativos(emisor)
    minimo    = calcular_minimo(tokens)
    con_alias = usa_alias(emisor)
    for termino in nombres:
        if len(resultados) >= 6: break
        queries = [
            (f'{termino} (downgrade OR upgrade OR outlook OR rating OR Moody OR Fitch OR "S&P" OR perspectiva OR calificacion OR bonos OR deuda OR sindicado)', True),
            (f'{termino} (finanzas OR resultados OR bolsa OR BVL OR SMV OR inversión OR accion OR utilidad)', False),
            (termino, False),
        ]
        for query, es_credit in queries:
            if len(resultados) >= 6: break
            try:
                url = f"https://news.google.com/rss/search?q={urllib.parse.quote(query)}&hl=es-419&gl=PE&ceid=PE:es-419&tbs=qdr:w,sbd:1"
                req = urllib.request.Request(url, headers=HEADERS)
                with urllib.request.urlopen(req, context=contexto, timeout=10) as res:
                    data = res.read().decode('utf-8', errors='ignore')
                for item in re.findall(r'<item>([\s\S]*?)</item>', data)[:6]:
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
                    vistos.add(titulo)
                    es_alerta = bool(re.search('|'.join(PALABRAS_CREDITICIAS), titulo, re.IGNORECASE))
                    resultados.append({
                        "titulo": titulo, "fuente": fuente,
                        "link": link, "fecha": formatear_fecha(fecha),
                        "alerta": es_alerta or es_credit
                    })
            except Exception:
                pass
    resultados.sort(key=lambda x: (0 if x["alerta"] else 1))
    return resultados[:6]

# ============================================================
# 11. CONSOLIDAR Y ORDENAR NOTICIAS
# ============================================================
def obtener_noticias(emisor):
    todas = (
        buscar_en_cache_prioritarias(emisor) +
        buscar_bloomberg_linea(emisor) +
        buscar_google_news(emisor)
    )
    vistos = set()
    resultado = []
    for n in todas:
        key = n["titulo"][:60].lower().strip()
        if key not in vistos:
            vistos.add(key)
            n["score"] = calcular_score(n["titulo"])
            resultado.append(n)
        if len(resultado) >= 6: break
    resultado.sort(key=lambda x: -(5 if x.get("alerta") else x.get("score", 0)))
    return resultado[:6]

# ============================================================
# 12. HELPERS HTML
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
    if n["alerta"]:                        return "bg-red-500"
    if "bvl"       in n["fuente"].lower(): return "bg-blue-500"
    if "smv"       in n["fuente"].lower(): return "bg-indigo-500"
    if "bloomberg" in n["fuente"].lower(): return "bg-orange-500"
    return "bg-gray-500"

def render_cards(emisores_con_noticias, seg):
    tc, bc, bgc = COLOR_SEG.get(seg, ("text-gray-400", "border-gray-600", "bg-gray-800"))
    cards = []
    for emisor, noticias in emisores_con_noticias:
        nombre_display = limpiar_nombre(emisor)
        tiene_alerta   = any(n["alerta"] for n in noticias)
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

        noticias_html = []
        for n in noticias:
            dc  = dot_color(n)
            ibg = "bg-red-950/30 border-red-500/40" if n["alerta"] else "bg-gray-800/40 border-gray-700/40"
            itx = "text-red-200" if n["alerta"] else "text-gray-300"
            fb  = badge_fuente(n["fuente"])
            noticias_html.append(f"""
              <div class="border rounded-lg p-2.5 {ibg}">
                <div class="flex gap-2 items-start">
                  <span class="mt-1.5 w-2 h-2 rounded-full shrink-0 {dc}"></span>
                  <div class="flex-1 min-w-0">
                    <a href="{html.escape(n['link'])}" target="_blank"
                       class="text-xs font-medium {itx} hover:text-blue-400 leading-snug line-clamp-3">
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
            <div class="flex flex-col gap-2">{''.join(noticias_html)}</div>
          </div>""")
    return ''.join(cards)

# ============================================================
# 13. EJECUCIÓN PRINCIPAL
# ============================================================
cargar_fuentes_prioritarias()
emisores_data = leer_emisores_excel("Emisores.xlsx")

if not emisores_data:
    print("ADVERTENCIA: No se leyeron emisores. Verifica Emisores.xlsx")
else:
    print(f"\n📋 {len(emisores_data)} emisores cargados desde Excel:")
    for e, s in emisores_data.items():
        v     = variantes_emisor(e)
        color = get_semaforo(e)
        sem   = f"[{color.upper()}]" if color else "[sin semáforo]"
        prior = "⭐" if es_emisor_prioritario(e) else ""
        print(f"   {limpiar_nombre(e):50s} → {v[0]:30s} {sem} {prior}")

segmentos = {}
for emisor, seg in emisores_data.items():
    segmentos.setdefault(seg, []).append(emisor)

segs_ordenados = sorted(segmentos.keys(), key=lambda s: ORDEN_SEG.index(s) if s in ORDEN_SEG else 99)

# ============================================================
# 14. PRE-CALCULAR NOTICIAS Y ORDENAR EMISORES
# ============================================================
print("\n🔍 Buscando noticias...")
resultados_por_segmento = {}

for seg in segs_ordenados:
    emisores_con_noticias = []
    for emisor in segmentos[seg]:
        nombre_display = limpiar_nombre(emisor)
        v = variantes_emisor(emisor)
        print(f"  ▶ {nombre_display:50s} → {v[0]}")
        noticias = obtener_noticias(emisor)
        if noticias:
            emisores_con_noticias.append((emisor, noticias))
            print(f"     ✓ {len(noticias)} noticias")
        else:
            print(f"     ↳ sin noticias, omitido")

    def sort_key_emisor(item):
        emisor, noticias = item
        return (
            0 if es_emisor_prioritario(emisor) else 1,
            0 if any(n.get("alerta") for n in noticias) else 1,
            -max((n.get("score", 0) for n in noticias), default=0)
        )

    emisores_con_noticias.sort(key=sort_key_emisor)

    if emisores_con_noticias:
        resultados_por_segmento[seg] = emisores_con_noticias
    else:
        print(f"  ⚠ Segmento '{seg}' sin noticias, omitido")

# ============================================================
# 15. CONSTRUCCIÓN HTML CON TABS
# ============================================================
total_emisores = sum(len(v) for v in resultados_por_segmento.values())

# Construir tabs nav
tabs_nav = f"""
    <button onclick="mostrarTab('todos')" id="tab-btn-todos"
      class="tab-btn px-4 py-2 rounded-lg text-sm font-semibold border transition-all
             bg-white/10 border-white/20 text-white">
      📋 Todos <span class="ml-1 text-[10px] opacity-60">({total_emisores})</span>
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

# Construir contenido de tabs
tabs_content = '<div id="tab-todos" class="tab-panel">'
for seg, items in resultados_por_segmento.items():
    tc = COLOR_SEG.get(seg, ("text-gray-400",))[0]
    tabs_content += f"""
    <div class="mb-12">
      <h2 class="text-lg font-bold mb-5 pb-2 border-b border-gray-800 {tc}">{html.escape(seg)}</h2>
      <div class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-5">{render_cards(items, seg)}</div>
    </div>"""
tabs_content += '</div>'

for seg, items in resultados_por_segmento.items():
    tc = COLOR_SEG.get(seg, ("text-gray-400",))[0]
    sid = seg.lower().replace(' ', '-')
    tabs_content += f"""
<div id="tab-{sid}" class="tab-panel hidden">
  <h2 class="text-lg font-bold mb-5 pb-2 border-b border-gray-800 {tc}">{html.escape(seg)}</h2>
  <div class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-5">{render_cards(items, seg)}</div>
</div>"""

# HTML final
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
        Fuentes: BVL · SMV · Bloomberg Línea · Google News &nbsp;·&nbsp; Última semana
      </p>
    </div>
    <div class="text-xs text-gray-500 font-mono">Actualizado: {fecha_reporte}</div>
  </header>

  <div class="flex flex-wrap gap-3 mb-6 text-xs">
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2 h-2 rounded-full bg-red-500"></span> Alerta crediticia
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2 h-2 rounded-full bg-blue-500"></span> BVL Oficial
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2 h-2 rounded-full bg-indigo-500"></span> SMV
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2 h-2 rounded-full bg-orange-500"></span> Bloomberg Línea
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2 h-2 rounded-full bg-gray-500"></span> Prensa general
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      🔴 Riesgo Alto &nbsp;·&nbsp; 🟡 Riesgo Moderado &nbsp;·&nbsp; 🟢 Riesgo Bajo
    </span>
  </div>

  <div class="flex flex-wrap gap-2 mb-8 border-b border-gray-800 pb-4">
    {tabs_nav}
  </div>

  {tabs_content if resultados_por_segmento else
   '<div class="bg-gray-900 border border-gray-800 rounded-xl p-12 text-center max-w-xl mx-auto my-12"><p class="text-emerald-400 font-medium text-lg mb-2">☕ Todo tranquilo</p><p class="text-gray-400 text-sm">No se encontraron noticias esta semana para ningún emisor.</p></div>'}

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
# 16. GUARDAR
# ============================================================
with open("index.html", "w", encoding="utf-8") as f:
    f.write(page)

print(f"\n✅ index.html generado — {len(emisores_data)} emisores procesados.")
