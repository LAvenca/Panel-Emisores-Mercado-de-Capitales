import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET
import html
import zipfile
import re

def leer_emisores_y_productos_robusto(ruta_archivo):
    mapping = {}
    try:
        with zipfile.ZipFile(ruta_archivo, 'r') as z:
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                with z.open('xl/sharedStrings.xml') as strings_file:
                    content = strings_file.read().decode('utf-8')
                    shared_strings = re.findall(r'<t[^>]*>(.*?)</t>', content)

            with z.open('xl/worksheets/sheet1.xml') as sheet:
                content = sheet.read().decode('utf-8')
                filas = re.findall(r'<row[^>]*>(.*?)</row>', content)
                
                for idx, fila in enumerate(filas):
                    celdas = re.findall(r'<c[^>]*>(.*?)</c>|<c[^>]*/>', fila)
                    if len(celdas) >= 2:
                        datos_fila = []
                        for celda in celdas[:2]:
                            v_match = re.search(r'<v>(.*?)</v>', celda)
                            if v_match:
                                val = v_match.group(1)
                                if 't="s"' in celda:
                                    try: texto = shared_strings[int(val)]
                                    except: texto = val
                                else:
                                    texto = val
                                texto = re.sub(r'<[^>]+>', '', texto).strip()
                                datos_fila.append(texto)
                            else:
                                datos_fila.append("")
                        
                        if len(datos_fila) >= 2:
                            emisor = datos_fila[0]
                            producto = datos_fila[1]
                            if idx == 0 and emisor.lower() in ['emisores', 'emisor', 'nombre', 'empresa']:
                                continue
                            if emisor and producto:
                                prod_normalizado = producto.strip().lower()
                                if 'fija' in prod_normalizado: mapping[emisor] = "Renta Fija"
                                elif 'variable' in prod_normalizado or 'accion' in prod_normalizado: mapping[emisor] = "Renta Variable"
                                elif 'fondo' in prod_normalizado: mapping[emisor] = "Fondos"
                                elif 'alterna' in prod_normalizado or 'alt' in prod_normalizado: mapping[emisor] = "Alternativos"
    except Exception as e:
        print("Cargando lista base de emisores...")
    return mapping

# Cargar emisores del portafolio
mapping_emisores = leer_emisores_y_productos_robusto("Emisores.xlsx")

# Clasificación forzada de seguridad para consistencia en el tablero
reglas_seguridad = {
    "aceros arequipa": "Renta Variable", "alicorp": "Renta Variable", "volcan": "Renta Variable",
    "inretail": "Renta Variable", "credicorp": "Renta Variable", "bcp": "Renta Variable",
    "ferreycorp": "Renta Variable", "minsur": "Renta Variable", "buenaventura": "Renta Variable",
    "intercorp": "Fondos", "fibra prime": "Alternativos", "luz del sur": "Renta Fija",
    "enel": "Renta Fija", "telefonica": "Renta Fija", "cerro verde": "Renta Variable",
    "banco gnb": "Renta Fija", "caja arequipa": "Renta Fija", "fossal": "Renta Variable"
}

if not mapping_emisores:
    mapping_emisores = {
        "Aceros Arequipa": "Renta Variable", "Alicorp": "Renta Variable", "Volcan": "Renta Variable",
        "Cerro Verde": "Renta Variable", "Banco GNB": "Renta Fija", "Caja Arequipa": "Renta Fija"
    }
else:
    for emisor in list(mapping_emisores.keys()):
        nombre_min = emisor.lower()
        for clave, prod_correcto in reglas_seguridad.items():
            if clave in nombre_min:
                mapping_emisores[emisor] = prod_correcto

orden_productos = ["Renta Fija", "Renta Variable", "Fondos", "Alternativos"]
datos_centralizados = {prod: {} for prod in orden_productos}
for emisor, prod in mapping_emisores.items():
    if prod in datos_centralizados:
        datos_centralizados[prod][emisor] = {"noticias": []}

# Cabeceras completas de simulación de navegador humano
HEADERS_NATIVOS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
    'Accept-Language': 'es-ES,es;q=0.9,en;q=0.8'
}

# Diccionario semántico de alertas crediticias y estructurales
patrones_criticos = [
    r'downgrade', r'upgrade', r'moody', r'fitch', r's&p', r'calificaci', r'clasificaci',
    r'perspectiva', r'rating', r'riesgo', r'sindicado', r'aumento de capital', r'hecho de importancia'
]
patron_riesgo = re.compile('|'.join(patrones_criticos), re.IGNORECASE)

# FUENTES PRIORITARIAS (Tus links obligatorios)
PAGINAS_PRIORITARIAS = [
    {"url": "https://www.bvl.com.pe/emisores/noticias-emisores", "fuente": "BVL Oficial"},
    {"url": "https://www.smv.gob.pe/SIMV/frm_hechosdeImportanciaDia?data=38C2EC33FA106691BB5B5039DACFDF50795D8EC3AF", "fuente": "SMV Diario"},
    {"url": "https://www.smv.gob.pe/SIMV/Frm_HechosDeImportancia?data=EBE76110FDC9EF5632D5100F5B0448927EBDAC2CF7", "fuente": "SMV Historial"},
    {"url": "https://www.moodys.com/entity/489500/overview", "fuente": "Moody's Radar"},
    {"url": "https://www.bloomberglinea.com/tags/las-noticias-del-dia/", "fuente": "Bloomberg Línea"}
]

links_globales_procesados = set()

# --- FASE 1: ESCANEO DIRECTO EN TUS FUENTES PRIORITARIAS ---
print("Fase 1: Extrayendo información directa de fuentes prioritarias...")
for portal in PAGINAS_PRIORITARIAS:
    try:
        req = urllib.request.Request(portal["url"], headers=HEADERS_NATIVOS)
        with urllib.request.urlopen(req, timeout=12) as response:
            html_puro = response.read().decode('utf-8', errors='ignore')
        
        texto_limpio = re.sub(r'<script[^>]*>([\s\S]*?)</script>|<style[^>]*>([\s\S]*?)</style>', '', html_puro)
        lineas = [re.sub(r'<[^>]+>', ' ', l).strip() for l in texto_limpio.split('\n') if l.strip()]
        
        for linea in lineas:
            if len(linea) < 15 or len(linea) > 230: continue
            
            for emisor, prod in mapping_emisores.items():
                if emisor.lower() in linea.lower() or (len(emisor) > 5 and emisor.lower()[:5] in linea.lower()):
                    es_prioritaria = bool(patron_riesgo.search(linea)) or "moody" in portal["fuente"].lower()
                    
                    if linea in links_globales_procesados: continue
                    links_globales_processed.add(linea)
                    
                    datos_centralizados[prod][emisor]["noticias"].append({
                        "titulo": linea,
                        "link": portal["url"],
                        "fuente": portal["fuente"],
                        "prioritaria": es_prioritaria
                    })
    except Exception as e:
        print(f"Portal diferido {portal['fuente']}")

# --- FASE 2: RASTREO ABIERTO EN LA RED (Google News Perú + México + USA Global) ---
print("Fase 2: Ejecutando rastreo complementario en la red global...")
for idx, (emisor, prod) in enumerate(mapping_emisores.items()):
    if prod not in datos_centralizados: continue
    
    limite_prensa = 2 if len(datos_centralizados[prod][emisor]["noticias"]) > 0 else 4
    
    emisor_encoded = urllib.parse.quote(emisor)
    urls_red = [
        f"https://news.google.com/rss/search?q={emisor_encoded}%20Peru&hl=es-419&gl=PE&ceid=PE:es-419",
        f"https://news.google.com/rss/search?q={emisor_encoded}%20Mexico&hl=es-419&gl=MX&ceid=MX:es-419",
        f"https://news.google.com/rss/search?q={emisor_encoded}%20(Moody%27s%20OR%20Downgrade%20OR%20Rating)&hl=es-419&gl=US&ceid=US:es-419"
    ]
    
    for url in urls_red:
        try:
            req = urllib.request.Request(url, headers=HEADERS_NATIVOS)
            with urllib.request.urlopen(req, timeout=8) as response:
                xml_data = response.read()
            root = ET.fromstring(xml_data)
            
            count = 0
            for item in root.findall('.//item'):
                if count >= limite_prensa: break
                titulo = item.find('title').text
                link = item.find('link').text
                fuente = item.find('source').text if item.find('source') is not None else "Red Internacional"
                
                if link in links_globales_procesados or titulo in links_globales_procesados: continue
                links_globales_procesados.add(link)
                links_globales_procesados.add(titulo)
                
                es_prioritaria = bool(patron_riesgo.search(titulo))
                
                datos_centralizados[prod][emisor]["noticias"].append({
                    "titulo": titulo,
                    "link": link,
                    "fuente": fuente,
                    "prioritaria": es_prioritaria
                })
                count += 1
        except:
            pass

# --- FASE 3: ORDENACIÓN DE ALERTA DE CRÉDITO Y FILTRADO ---
for prod in orden_productos:
    for emisor in datos_centralizados[prod]:
        datos_centralizados[prod][emisor]["noticias"].sort(key=lambda x: x["prioritaria"], reverse=True)
        datos_centralizados[prod][emisor]["noticias"] = datos_centralizados[prod][emisor]["noticias"][:5]

# --- FASE 4: MAQUETADO DE LA TERMINAL HTML ---
html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Monitoreo de Noticias del Portafolio</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-950 text-gray-100 min-h-screen font-sans">
    <div class="max-w-7xl mx-auto px-4 py-8">
        
        <header class="border-b border-gray-800 pb-6 mb-12 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
            <div>
                <h1 class="text-3xl font-extrabold bg-gradient-to-r from-blue-400 via-indigo-400 to-emerald-400 bg-clip-text text-transparent">
                    Monitoreo de Noticias del Portafolio
                </h1>
                <p class="text-gray-400 text-sm mt-1">Filtro de Crédito Avanzado • Fuentes Clave e Inteligencia Global Activos</p>
            </div>
            <div class="bg-red-950/40 border border-red-500/30 px-4 py-2 rounded-xl text-xs font-mono text-red-400 flex items-center gap-2">
                <span class="w-2 h-2 rounded-full bg-red-500 animate-pulse"></span> Escaneo Híbrido: Links Core + Red USA/Global
            </div>
        </header>
"""

total_visibles = 0
for producto in orden_productos:
    emisores_del_producto = datos_centralizados[producto]
    # CORREGIDO AQUÍ: Variable alineada con el flujo lógico
    contiene_noticias_activas = any(contenido["noticias"] for emisor, contenido in emisores_del_producto.items())
    if not contiene_noticias_activas: continue

    badge_color = "text-blue-400 border-blue-500/30 bg-blue-500/5"
    if producto == "Renta Variable": badge_color = "text-purple-400 border-purple-500/30 bg-purple-500/5"
    elif producto == "Fondos": badge_color = "text-emerald-400 border-emerald-500/30 bg-emerald-500/5"
    elif producto == "Alternativos": badge_color = "text-amber-400 border-amber-500/30 bg-amber-500/5"

    html_content += f"""
        <section class="mb-12">
            <div class="flex items-center gap-3 mb-6 border-b border-gray-800 pb-2">
                <h2 class="text-xl font-bold text-white">{producto}</h2>
                <span class="border text-xs px-2.5 py-0.5 rounded-md font-medium {badge_color}">Segmento</span>
            </div>
            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
    """

    for emisor, contenido in list(emisores_del_producto.items()):
        noticias = contenido["noticias"]
        if not noticias: continue
        total_visibles += 1
        
        tiene_alerta = any(n["prioritaria"] for n in noticias)
        borde_caja = "border-red-500/50 bg-gradient-to-b from-gray-900 to-red-950/20 shadow-red-950/30 shadow-lg" if tiene_alerta else "border-gray-800 bg-gray-900 shadow-lg"

        html_content += f"""
                <div class="rounded-xl border p-5 flex flex-col justify-between hover:border-gray-700 transition duration-300 {borde_caja}">
                    <div>
                        <div class="pb-2 mb-3 border-b border-gray-800/60 flex justify-between items-center">
                            <h3 class="text-sm font-bold text-gray-100 tracking-wide uppercase">{html.escape(emisor)}</h3>
                            { '<span class="bg-red-500/20 text-red-400 border border-red-500/30 text-[9px] px-1.5 py-0.5 rounded uppercase font-bold tracking-wider animate-pulse">Rating Alert</span>' if tiene_alerta else '' }
                        </div>
                        <div class="space-y-2">
        """
        for n in noticias:
            bg_item = "bg-red-500/10 border-red-500 text-red-200 font-medium" if n["prioritaria"] else "bg-gray-850 border-gray-700 text-gray-300"
            html_content += f"""
                            <div class="border-l-2 p-2 rounded-r text-xs {bg_item}">
                                <a href="{n['link']}" target="_blank" class="hover:text-blue-400 line-clamp-3">{html.escape(n['titulo'])}</a>
                                <div class="flex justify-between text-[9px] text-gray-500 mt-1 font-mono">
                                    <span>Origen: {html.escape(n['fuente'])}</span>
                                </div>
                            </div>
            """
        html_content += """</div></div></div>"""
    html_content += """</div></section>"""

if total_visibles == 0:
    html_content += """
        <div class="bg-gray-900 border border-gray-800 rounded-xl p-12 text-center max-w-xl mx-auto my-12">
            <p class="text-emerald-400 font-medium text-lg mb-2">☕ Mercado bajo monitoreo</p>
            <p class="text-gray-400 text-sm">No se reportan eventos para los emisores activos del portafolio en este instante.</p>
        </div>
    """

html_content += """
        <footer class="mt-16 pt-6 border-t border-gray-900 text-center text-xs text-gray-600">
            Filtro de Crédito Customizado Multiregional • BVL, SMV, Moody's, Bloomberg Línea y Google Global.
        </footer>
    </div>
</body>
</html>
