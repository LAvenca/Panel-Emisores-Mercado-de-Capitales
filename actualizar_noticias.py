import urllib.request
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
        print(f"Nota: No se pudo procesar estructura extendida del Excel ({e})")
    return mapping

# Cargar configuración inicial
mapping_emisores = leer_emisores_y_productos_robusto("Emisores.xlsx")

# Base de datos manual de clasificación forzada para asegurar consistencia de instrumentos
reglas_seguridad = {
    "aceros arequipa": "Renta Variable",
    "alicorp": "Renta Variable",
    "volcan": "Renta Variable",
    "inretail": "Renta Variable",
    "credicorp": "Renta Variable",
    "bcp": "Renta Variable",
    "ferreycorp": "Renta Variable",
    "minsur": "Renta Variable",
    "buenaventura": "Renta Variable",
    "intercorp": "Fondos",
    "fibra prime": "Alternativos",
    "luz del sur": "Renta Fija",
    "enel": "Renta Fija",
    "telefonica": "Renta Fija"
}

if not mapping_emisores:
    mapping_emisores = {
        "Aceros Arequipa": "Renta Variable",
        "Alicorp": "Renta Variable",
        "Volcan": "Renta Variable",
        "Pacífico Seguros": "Renta Fija",
        "Rímac Seguros": "Renta Fija"
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
        datos_centralizados[prod][emisor] = {"hechos": [], "noticias": []}

# Patrón semántico de alertas de calificación crediticia (Upgrades, Downgrades, Perspectivas)
patron_riesgo = re.compile(r'(downgrade|upgrade|calificacion|clasificacion|clasificadora|perspectiva|rating|bajada|subida|investment grade|grado de inversion)', re.IGNORECASE)

# Descargar Hechos Relevantes desde la SMV
print("Buscando alertas de clasificación en la SMV...")
url_smv = "https://www.smv.gob.pe/Frm_HechosRelevantesRSS?id=0"
try:
    req_smv = urllib.request.Request(url_smv, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req_smv, timeout=15) as response:
        xml_raw = response.read().decode('utf-8', errors='ignore')
    
    xml_saneado = re.sub(r'&(?! [a-zA-Z0-9#]+;)', '&amp;', xml_raw)
    root_smv = ET.fromstring(xml_saneado)
    
    for item in root_smv.findall('.//item'):
        titulo = item.find('title').text if item.find('title') is not None else ""
        link = item.find('link').text if item.find('link') is not None else ""
        fecha = item.find('pubDate').text if item.find('pubDate') is not None else ""
        descripcion = item.find('description').text if item.find('description') is not None else ""
        texto_analizar = (titulo + " " + descripcion).lower()
        
        for emisor, prod in mapping_emisores.items():
            if prod in datos_centralizados and (emisor.lower() in texto_analizar or (len(emisor) > 5 and emisor.lower()[:5] in texto_analizar)):
                es_prioritaria = bool(patron_riesgo.search(texto_analizar))
                
                datos_centralizados[prod][emisor]["hechos"].append({
                    "titulo": titulo.replace("Hecho Relevante -", "").strip(),
                    "link": link,
                    "fecha": fecha[:16],
                    "prioritaria": es_prioritaria
                })
except Exception as e:
    print(f"Aviso SMV: {e}")

# Descargar noticias desde Google News
print("Buscando noticias en Google News...")
for idx, (emisor, prod) in enumerate(mapping_emisores.items()):
    if prod not in datos_centralizados: continue
    if idx > 25: continue
    
    query = urllib.parse.quote(f"{emisor} Peru")
    url = f"https://news.google.com/rss/search?q={query}&hl=es-419&gl=PE&ceid=PE:es-419"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=8) as response:
            xml_data = response.read()
        root = ET.fromstring(xml_data)
        
        for item in root.findall('.//item')[:4]:
            titulo = item.find('title').text
            link = item.find('link').text
            fecha = item.find('pubDate').text
            fuente = item.find('source').text if item.find('source') is not None else "Prensa"
            
            es_prioritaria = bool(patron_riesgo.search(titulo))
            
            datos_centralizados[prod][emisor]["noticias"].append({
                "titulo": titulo,
                "link": link,
                "fecha": fecha[:16],
                "fuente": fuente,
                "prioritaria": es_prioritaria
            })
    except Exception as e:
        pass

# ORDENACIÓN INTERNA POR RELEVANCIA DE CRÉDITO
for prod in orden_productos:
    for emisor in datos_centralizados[prod]:
        datos_centralizados[prod][emisor]["hechos"].sort(key=lambda x: x["prioritaria"], reverse=True)
        datos_centralizados[prod][emisor]["noticias"].sort(key=lambda x: x["prioritaria"], reverse=True)

# 5. MAQUETADO HTML
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
                <p class="text-gray-400 text-sm mt-1">Filtro de Crédito con prioridad en Calificaciones de Riesgo e Instrumentos</p>
            </div>
            <div class="bg-red-950/40 border border-red-500/30 px-4 py-2 rounded-xl text-xs font-mono text-red-400 flex items-center gap-2 animate-pulse">
                <span class="w-2 h-2 rounded-full bg-red-500"></span> Alertas de Calificación Priorizadas
            </div>
        </header>
"""

total_emisores_visibles = 0

for producto in orden_productos:
    emisores_del_producto = datos_centralizados[producto]
    
    # CRITERIO DE EXCLUSIÓN COMPLETA: Si no hay noticias hoy, no sale la categoría
    contiene_noticias_activas = False
    for emisor, contenido in list(emisores_del_producto.items()):
        if contenido["hechos"] or contenido["noticias"]:
            contiene_noticias_activas = True
            break
            
    if not contiene_noticias_activas:
        continue

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
        hechos = contenido["hechos"]
        noticias = contenido["noticias"]
        
        # CONDICIÓN ESTRICTA: Si la empresa no trae información hoy, NO sale en el tablero
        if not hechos and not noticias:
            continue
            
        total_emisores_visibles += 1

        tiene_alerta_critica = any(h["prioritaria"] for h in hechos) or any(n["prioritaria"] for n in noticias)
        borde_caja = "border-red-500/40 bg-gradient-to-b from-gray-900 to-red-950/10 shadow-red-950/20" if tiene_alerta_critica else "border-gray-800 bg-gray-900 shadow-lg"

        html_content += f"""
                <div class="rounded-xl border p-5 flex flex-col justify-between hover:border-gray-700 transition duration-300 {borde_caja}">
                    <div>
                        <div class="pb-2 mb-3 border-b border-gray-800/60 flex justify-between items-center">
                            <h3 class="text-sm font-bold text-gray-100 tracking-wide uppercase">{html.escape(emisor)}</h3>
                            { '<span class="bg-red-500/10 text-red-400 border border-red-500/20 text-[9px] px-1.5 py-0.5 rounded uppercase font-bold tracking-wider animate-pulse">Rating Alert</span>' if tiene_alerta_critica else '' }
                        </div>
        """

        if hechos:
            html_content += """<div class="mb-3"><h4 class="text-[10px] font-bold text-amber-500 uppercase tracking-wider mb-1">🚨 SMV / Hechos</h4><div class="space-y-1.5">"""
            for h in hechos:
                bg_item = "bg-red-500/10 border-red-500 font-semibold" if h["prioritaria"] else "bg-amber-500/5 border-amber-500"
                html_content += f"""<div class="border-l-2 p-2 rounded-r text-xs {bg_item}">
                                    <a href="{h['link']}" target="_blank" class="text-gray-200 hover:text-blue-400 line-clamp-2">{html.escape(h['titulo'])}</a>
                                </div>"""
            html_content += """</div></div>"""

        if noticias:
            html_content += """<div><h4 class="text-[10px] font-bold text-emerald-400 uppercase tracking-wider mb-1">📰 Prensa</h4><div class="space-y-1.5">"""
            for n in noticias:
                bg_item = "bg-red-500/10 border-red-500 font-semibold" if n["prioritaria"] else "bg-emerald-500/5 border-emerald-500"
                html_content += f"""<div class="border-l-2 p-2 rounded-r text-xs {bg_item}">
                                    <a href="{n['link']}" target="_blank" class="text-gray-200 hover:text-blue-400 line-clamp-2">{html.escape(n['titulo'])}</a>
                                </div>"""
            html_content += """</div></div>"""

        html_content += """</div></div>"""

    html_content += """</div></section>"""

if total_emisores_visibles == 0:
    html_content += """
        <div class="bg-gray-900 border border-gray-800 rounded-xl p-12 text-center max-w-xl mx-auto my-12">
            <p class="text-emerald-400 font-medium text-lg mb-2">☕ Todo bajo control</p>
            <p class="text-gray-400 text-sm">No se reportan hechos de importancia ni cambios de perspectivas para el portafolio en este ciclo.</p>
        </div>
    """

html_content += """
        <footer class="mt-16 pt-6 border-t border-gray-900 text-center text-xs text-gray-600">
            Filtro regulatorio segmentado por categorías de inversión con priorización de riesgo.
        </footer>
    </div>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
print("¡Fichero index.html generado con éxito!")
