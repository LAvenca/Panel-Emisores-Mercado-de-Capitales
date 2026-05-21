import urllib.request
import urllib.parse
import re
import ssl
import html
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
# 2. CONFIGURACIÓN GLOBAL
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

LIMITE_FECHA = datetime.now() - timedelta(days=30)   # Máximo 1 mes
fecha_reporte = datetime.now().strftime("%d/%m/%Y %H:%M")

# ============================================================
# 3. UTILIDADES DE FECHA
# ============================================================
def parsear_fecha(fecha_raw):
    """Intenta parsear fecha RSS. Devuelve datetime o None."""
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
    # Fallback: extraer día mes año
    m = re.search(r'(\d{1,2})\s+(\w{3})\s+(\d{4})', fecha_raw)
    if m:
        try:
            return datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%d %b %Y")
        except Exception:
            pass
    return None

def formatear_fecha(fecha_raw):
    """Devuelve string legible en español."""
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
    """True si la noticia es del último mes."""
    dt = parsear_fecha(fecha_raw)
    if not dt:
        return True   # Si no podemos parsear, la dejamos pasar
    return dt >= LIMITE_FECHA

# ============================================================
# 4. SCRAPING DE FUENTES PRIORITARIAS (BVL, SMV, Bloomberg Línea)
# ============================================================
FUENTES_PRIORITARIAS = [
    {
        "nombre": "BVL Oficial",
        "url": "https://www.bvl.com.pe/emisores/noticias-emisores",
        "tipo": "html"
    },
    {
        "nombre": "SMV Diario",
        "url": "https://www.smv.gob.pe/SIMV/frm_hechosdeImportanciaDia?data=38C2EC33FA106691BB5B5039DACFDF50795D8EC3AF",
        "tipo": "html"
    },
]

# Cache global de noticias por fuente prioritaria
# { texto_linea: {"titulo": ..., "fuente": ..., "link": ..., "fecha": ...} }
cache_prioritarias = []

def cargar_fuentes_prioritarias():
    print("  Cargando fuentes prioritarias (BVL, SMV)...")
    for fuente in FUENTES_PRIORITARIAS:
        try:
            req = urllib.request.Request(fuente["url"], headers=HEADERS)
            with urllib.request.urlopen(req, context=contexto, timeout=10) as res:
                raw = res.read().decode('utf-8', errors='ignore')

            # Limpiar scripts y estilos
            raw = re.sub(r'<script[\s\S]*?</script>', '', raw, flags=re.IGNORECASE)
            raw = re.sub(r'<style[\s\S]*?</style>', '', raw, flags=re.IGNORECASE)

            # Extraer links con texto
            links = re.findall(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>([\s\S]*?)</a>', raw, re.IGNORECASE)
            for href, texto in links:
                texto_limpio = re.sub(r'<[^>]+>', ' ', texto).strip()
                texto_limpio = re.sub(r'\s+', ' ', texto_limpio)
                if len(texto_limpio) < 15 or len(texto_limpio) > 300:
                    continue
                if not href.startswith('http'):
                    base = re.match(r'https?://[^/]+', fuente["url"])
                    href = (base.group(0) if base else "") + "/" + href.lstrip("/")

                cache_prioritarias.append({
                    "titulo": texto_limpio,
                    "fuente": fuente["nombre"],
                    "link": href,
                    "fecha": "Hoy",
                    "alerta": bool(re.search('|'.join(PALABRAS_CREDITICIAS), texto_limpio, re.IGNORECASE))
                })

            # También extraer líneas de texto puro (para SMV que usa tablas)
            lineas = [re.sub(r'<[^>]+>', ' ', l).strip() for l in raw.split('\n')]
            for linea in lineas:
                linea = re.sub(r'\s+', ' ', linea).strip()
                if len(linea) < 20 or len(linea) > 300:
                    continue
                cache_prioritarias.append({
                    "titulo": linea,
                    "fuente": fuente["nombre"],
                    "link": fuente["url"],
                    "fecha": "Hoy",
                    "alerta": bool(re.search('|'.join(PALABRAS_CREDITICIAS), linea, re.IGNORECASE))
                })

            print(f"    ✓ {fuente['nombre']}: {len(cache_prioritarias)} entradas cargadas")
        except Exception as e:
            print(f"    ✗ {fuente['nombre']} no disponible: {e}")

def buscar_bloomberg_linea(emisor):
    """Busca en Bloomberg Línea vía Google News RSS."""
    resultados = []
    try:
        q = urllib.parse.quote(f'"{emisor}" site:bloomberglinea.com')
        url = f"https://news.google.com/rss/search?q={q}&hl=es-419&gl=PE&ceid=PE:es-419&tbs=qdr:m,sbd:1"
        req = urllib.request.Request(url, headers=HEADERS)
        with urllib.request.urlopen(req, context=contexto, timeout=10) as res:
            data = res.read().decode('utf-8', errors='ignore')

        items = re.findall(r'<item>([\s\S]*?)</item>', data)
        for item in items[:5]:
            t = re.search(r'<title><!\[CDATA\[(.*?)\]\]></title>|<title>(.*?)</title>', item)
            l = re.search(r'<link>(.*?)</link>', item)
            d = re.search(r'<pubDate>(.*?)</pubDate>', item)

            titulo = (t.group(1) or t.group(2) or "").strip() if t else ""
            link   = l.group(1).strip() if l else "#"
            fecha  = d.group(1).strip() if d else ""

            if not titulo or not es_reciente(fecha):
                continue

            resultados.append({
                "titulo": titulo,
                "fuente": "Bloomberg Línea",
                "link": link,
                "fecha": formatear_fecha(fecha),
                "alerta": bool(re.search('|'.join(PALABRAS_CREDITICIAS), titulo, re.IGNORECASE))
            })
    except Exception:
        pass
    return resultados

def buscar_en_cache_prioritarias(emisor):
    """Busca el nombre del emisor en el cache de BVL/SMV."""
    resultados = []
    raiz = emisor.split()[0].lower().replace(",", "").replace(".", "")
    if len(raiz) <= 3 and len(emisor.split()) > 1:
        raiz = emisor.split()[1].lower()

    vistos = set()
    for entrada in cache_prioritarias:
        if raiz in entrada["titulo"].lower():
            key = entrada["titulo"][:60]
            if key not in vistos:
                vistos.add(key)
                resultados.append(entrada)
    return resultados

# ============================================================
# 5. BÚSQUEDA PRINCIPAL (Google News RSS)
# ============================================================
def buscar_google_news(emisor):
    resultados = []
    queries = [
        # Alta prioridad: alertas crediticias
        (f'"{emisor}" (downgrade OR upgrade OR outlook OR rating OR Moody OR Fitch OR "S&P" OR perspectiva OR calificacion OR bonos OR deuda)', True),
        # General reciente
        (f'"{emisor}" (finanzas OR resultados OR bolsa OR BVL OR SMV OR inversión)', False),
    ]

    vistos = set()
    for query, es_credit in queries:
        try:
            q_enc = urllib.parse.quote(query)
            url = (
                f"https://news.google.com/rss/search?"
                f"q={q_enc}&hl=es-419&gl=PE&ceid=PE:es-419&tbs=qdr:m,sbd:1"
            )
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, context=contexto, timeout=10) as res:
                data = res.read().decode('utf-8', errors='ignore')

            items = re.findall(r'<item>([\s\S]*?)</item>', data)
            for item in items[:6]:
                t = re.search(r'<title><!\[CDATA\[(.*?)\]\]></title>|<title>(.*?)</title>', item)
                l = re.search(r'<link>(.*?)</link>', item)
                d = re.search(r'<pubDate>(.*?)</pubDate>', item)
                s = re.search(r'<source[^>]*>(.*?)</source>', item)

                titulo = (t.group(1) or t.group(2) or "").strip() if t else ""
                link   = l.group(1).strip() if l else "#"
                fecha  = d.group(1).strip() if d else ""
                fuente = (s.group(1) or "Prensa").strip() if s else "Prensa"

                if not titulo or titulo in vistos:
                    continue
                if not es_reciente(fecha):
                    continue

                vistos.add(titulo)
                es_alerta = bool(re.search('|'.join(PALABRAS_CREDITICIAS), titulo, re.IGNORECASE)) or es_credit

                resultados.append({
                    "titulo": titulo,
                    "fuente": fuente,
                    "link": link,
                    "fecha": formatear_fecha(fecha),
                    "alerta": es_alerta
                })
        except Exception:
            pass

    return resultados

# ============================================================
# 6. CONSOLIDAR NOTICIAS POR EMISOR
# ============================================================
def obtener_noticias(emisor):
    """
    Orden de prioridad:
    1. BVL / SMV  (fuente oficial)
    2. Bloomberg Línea
    3. Google News general
    Todo filtrado al último mes y deduplicado.
    """
    noticias_bvl_smv  = buscar_en_cache_prioritarias(emisor)
    noticias_bloomberg = buscar_bloomberg_linea(emisor)
    noticias_google    = buscar_google_news(emisor)

    # Combinar respetando prioridad
    todas = noticias_bvl_smv + noticias_bloomberg + noticias_google

    # Deduplicar por título (primeros 60 chars)
    vistos = set()
    resultado = []
    for n in todas:
        key = n["titulo"][:60].lower().strip()
        if key not in vistos:
            vistos.add(key)
            resultado.append(n)
        if len(resultado) >= 6:
            break

    # Ordenar: alertas primero, luego el resto
    resultado.sort(key=lambda x: (0 if x["alerta"] else 1))
    return resultado[:6]

# ============================================================
# 7. EJECUTAR CARGA INICIAL
# ============================================================
cargar_fuentes_prioritarias()
emisores_data = leer_emisores_excel("Emisores.xlsx")

if not emisores_data:
    print("ADVERTENCIA: No se leyeron emisores. Verifica Emisores.xlsx")

# ============================================================
# 8. CONSTRUCCIÓN DEL HTML
# ============================================================
ORDEN_SEG  = ["Renta Fija", "Renta Variable", "Fondos", "Alternativos"]
COLOR_SEG  = {
    "Renta Fija":     ("text-blue-400",   "border-blue-500/30",   "bg-blue-500/10"),
    "Renta Variable": ("text-purple-400", "border-purple-500/30", "bg-purple-500/10"),
    "Fondos":         ("text-emerald-400","border-emerald-500/30","bg-emerald-500/10"),
    "Alternativos":   ("text-amber-400",  "border-amber-500/30",  "bg-amber-500/10"),
}

# Badge de fuente
def badge_fuente(fuente):
    f = fuente.lower()
    if "bvl" in f:
        return "bg-blue-900/40 text-blue-300 border-blue-600/40"
    if "smv" in f:
        return "bg-indigo-900/40 text-indigo-300 border-indigo-600/40"
    if "bloomberg" in f:
        return "bg-orange-900/40 text-orange-300 border-orange-600/40"
    return "bg-gray-800 text-gray-400 border-gray-700"

# Agrupar por segmento
segmentos = {}
for emisor, seg in emisores_data.items():
    segmentos.setdefault(seg, []).append(emisor)

segs_ordenados = sorted(segmentos.keys(), key=lambda s: ORDEN_SEG.index(s) if s in ORDEN_SEG else 99)

html_parts = []
html_parts.append(f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Monitor Crediticio — Portafolio</title>
  <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-950 text-gray-100 min-h-screen font-sans">
<div class="max-w-7xl mx-auto px-4 py-10">

  <!-- ENCABEZADO -->
  <header class="mb-10 border-b border-gray-800 pb-6 flex flex-col md:flex-row md:justify-between md:items-end gap-4">
    <div>
      <h1 class="text-3xl font-extrabold bg-gradient-to-r from-blue-400 via-indigo-400 to-emerald-400 bg-clip-text text-transparent">
        Monitor Crediticio del Portafolio
      </h1>
      <p class="text-gray-400 text-sm mt-1">
        Fuentes: BVL · SMV · Bloomberg Línea · Google News &nbsp;·&nbsp; Noticias del último mes
      </p>
    </div>
    <div class="text-right text-xs text-gray-500 font-mono">
      Actualizado: {fecha_reporte}
    </div>
  </header>

  <!-- LEYENDA -->
  <div class="flex flex-wrap gap-4 mb-8 text-xs">
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2.5 h-2.5 rounded-full bg-red-500"></span> Alerta crediticia (rating / downgrade / outlook)
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2.5 h-2.5 rounded-full bg-blue-500"></span> BVL Oficial
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2.5 h-2.5 rounded-full bg-indigo-500"></span> SMV
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2.5 h-2.5 rounded-full bg-orange-500"></span> Bloomberg Línea
    </span>
    <span class="flex items-center gap-1.5 bg-gray-900 border border-gray-800 px-3 py-1.5 rounded-lg">
      <span class="w-2.5 h-2.5 rounded-full bg-gray-500"></span> Prensa general
    </span>
  </div>
""")

for seg in segs_ordenados:
    col = COLOR_SEG.get(seg, ("text-gray-400", "border-gray-600", "bg-gray-800"))
    tc, bc, bgc = col

    html_parts.append(f"""
  <!-- SEGMENTO: {html.escape(seg)} -->
  <div class="mb-12">
    <h2 class="text-lg font-bold mb-5 pb-2 border-b border-gray-800 {tc}">
      {html.escape(seg)}
    </h2>
    <div class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-5">
""")

    for emisor in segmentos[seg]:
        print(f"  ▶ {emisor} ({seg})")
        noticias    = obtener_noticias(emisor)
        tiene_alerta = any(n["alerta"] for n in noticias)

        card_border = "border-red-500/50 shadow-red-950/30 shadow-lg" if tiene_alerta else "border-gray-800"
        card_bg     = "bg-gradient-to-b from-gray-900 to-red-950/15" if tiene_alerta else "bg-gray-900"

        html_parts.append(f"""
      <div class="rounded-xl border p-5 flex flex-col gap-3 {card_border} {card_bg}">

        <!-- Cabecera tarjeta -->
        <div class="flex justify-between items-center border-b border-gray-800/70 pb-2.5">
          <h3 class="text-sm font-bold uppercase tracking-wide">{html.escape(emisor)}</h3>
          <div class="flex items-center gap-1.5">
            <span class="text-[10px] border px-2 py-0.5 rounded {tc} {bc} {bgc}">{html.escape(seg)}</span>
            {'<span class="text-[9px] bg-red-500/20 border border-red-500/30 text-red-400 px-2 py-0.5 rounded font-bold animate-pulse">⚠ Rating Alert</span>' if tiene_alerta else ""}
          </div>
        </div>

        <!-- Noticias -->
        <div class="flex flex-col gap-2">
""")

        if noticias:
            for n in noticias:
                dot   = "bg-red-500" if n["alerta"] else (
                    "bg-blue-500"   if "bvl"       in n["fuente"].lower() else
                    "bg-indigo-500" if "smv"       in n["fuente"].lower() else
                    "bg-orange-500" if "bloomberg" in n["fuente"].lower() else
                    "bg-gray-500"
                )
                ibg   = "bg-red-950/30 border-red-500/40" if n["alerta"] else "bg-gray-800/40 border-gray-700/40"
                itxt  = "text-red-200"  if n["alerta"] else "text-gray-300"
                fbadge = badge_fuente(n["fuente"])

                html_parts.append(f"""
          <div class="border rounded-lg p-2.5 {ibg}">
            <div class="flex gap-2 items-start">
              <span class="mt-1.5 w-2 h-2 rounded-full shrink-0 {dot}"></span>
              <div class="flex-1 min-w-0">
                <a href="{html.escape(n['link'])}" target="_blank"
                   class="text-xs font-medium {itxt} hover:text-blue-400 leading-snug line-clamp-3">
                  {html.escape(n['titulo'])}
                </a>
                <div class="flex justify-between items-center mt-1.5 gap-2">
                  <span class="text-[10px] border px-1.5 py-0.5 rounded font-mono {fbadge}">
                    {html.escape(n['fuente'])}
                  </span>
                  <span class="text-[10px] text-gray-500 font-mono shrink-0">{html.escape(n['fecha'])}</span>
                </div>
              </div>
            </div>
          </div>
""")
        else:
            html_parts.append("""
          <p class="text-xs text-gray-600 italic py-2">Sin noticias recientes este mes.</p>
""")

        html_parts.append("""
        </div><!-- fin noticias -->
      </div><!-- fin card -->
""")

    html_parts.append("""
    </div><!-- fin grid -->
  </div><!-- fin segmento -->
""")

html_parts.append(f"""
  <footer class="mt-12 pt-6 border-t border-gray-800 text-center text-xs text-gray-600">
    Monitor Crediticio · BVL · SMV · Bloomberg Línea · Google News · Solo fuentes públicas
  </footer>

</div>
</body>
</html>
""")

# ============================================================
# 9. GUARDAR
# ============================================================
with open("index.html", "w", encoding="utf-8") as f:
    f.write("".join(html_parts))

print(f"\n✅ index.html generado — {len(emisores_data)} emisores procesados.")
