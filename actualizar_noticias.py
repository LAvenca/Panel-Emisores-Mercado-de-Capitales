import urllib.request
import xml.etree.ElementTree as ET
import html
import zipfile
import re

def leer_emisores_y_productos_excel(ruta_archivo):
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
                
                for index, fila in enumerate(filas):
                    celdas = re.findall(r'<c[^>]*>(.*?)</c>', fila)
                    if len(celdas) >= 2:
                        # Extraer Texto de Columna A (Emisor)
                        c1 = celdas[0]
                        v1_match = re.search(r'<v>(.*?)</v>', c1)
                        emisor = ""
                        if v1_match:
                            val1 = v1_match.group(1)
                            emisor = shared_strings[int(val1)] if 't="s"' in c1 and shared_strings else val1
                            emisor = re.sub(r'<[^>]+>', '', emisor).strip()

                        # Extraer Texto de Columna B (Producto)
                        c2 = celdas[1]
                        v2_match = re.search(r'<v>(.*?)</v>', c2)
                        producto = ""
                        if v2_match:
                            val2 = v2_match.group(1)
                            producto = shared_strings[int(val2)] if 't="s"' in c2 and shared_strings else val2
                            producto = re.sub(r'<[^>]+>', '', producto).strip()

                        # Omitir cabecera
                        if index == 0 and emisor.lower() in ['emisores', 'emisor', 'nombre']:
                            continue

                        if emisor and producto:
                            mapping[emisor] = producto
    except Exception as e:
        print(f"Error al leer Excel: {e}")
    return mapping

# 1. Cargar mapeo del Excel
mapping_emisores = leer_emisores_y_productos_excel("Emisores.xlsx")

# Lista de respaldo si el Excel está vacío
if not mapping_emisores:
    mapping_emisores = {
        "Pacífico Seguros": "Renta Fija",
        "Rímac Seguros": "Renta Fija",
        "Interseguro": "Renta Variable",
        "Credicorp": "Renta Variable",
        "Fibra Prime": "Alternativos"
    }

print(f"Mapeo cargado: {mapping_emisores}")

# Orden estricto solicitado
orden_productos = ["Renta Fija", "Renta Variable", "Fondos", "Alternativos"]

# Inicializar estructura centralizada organizada por Producto -> Emisor
datos_centralizados = {prod: {} for prod in orden_productos}
for emisor, prod in mapping_emisores.items():
    if prod in datos_centralizados:
        datos_centralizados[prod][emisor] = {"hechos": [], "noticias": []}

# 2. SECCIÓN SMV: Descargar e integrar Hechos Relevantes
print("Consultando Hechos Relevantes en la SMV...")
url_smv = "https://www.smv.gob.pe/Frm_HechosRelevantesRSS?id=0"
try:
    req_smv = urllib.request.Request(url_smv, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req_smv, timeout=15) as response:
        xml_smv = response.read()
    root_smv = ET.fromstring(xml_smv)
    
    for item in root_smv.findall('.//item'):
        titulo = item.find('title').text if item.find('title') is not None else ""
        link = item.find('link').text if item.find('link') is not None else ""
        fecha = item.find('pubDate').text if item.find('pubDate') is not None else ""
        descripcion = item.find('description').text if item.find('description') is not None else ""
        texto_analizar = (titulo + " " + descripcion).lower()
        
        for emisor, prod in mapping_emisores.items():
            if prod in datos_centralizados and (emisor.lower() in texto_analizar or (len(emisor) > 5 and emisor.lower()[:5] in texto_analizar)):
                datos_centralizados[prod][emisor]["hechos"].append({
                    "titulo": titulo.replace("Hecho Relevante -", "").strip(),
                    "link": link,
                    "fecha": fecha[:16]
                })
                break
except Exception as e:
    print(f"Nota: No se pudo conectar a la SMV ({e}).")

# 3. SECCIÓN GOOGLE NEWS: Descargar noticias de prensa
print("Consultando Google News...")
for emisor, prod in mapping_emisores.items():
    if prod not in datos_centralizados: continue
    query = urllib.parse.quote(f"{emisor} Peru")
    url = f"https://news.google.com/rss/search?q={query}&hl=es-419&gl=PE&ceid=PE:es-419"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=15) as response:
            xml_data = response.read()
        root = ET.fromstring(xml_data)
        count = 0
        for item in root.findall('.//item'):
            if count >= 3: break
            titulo = item.find('title').text
            link = item.find('link').text
            fecha = item.find('pubDate').text
            fuente = item.find('source').text if item.find('source') is not None else "Prensa"
            
            datos_centralizados[prod][emisor]["noticias"].append({
                "titulo": titulo,
                "link": link,
                "fecha": fecha[:16],
                "fuente": fuente
            })
            count += 1
    except Exception as e:
        print(f"Error en noticias para {emisor}: {e}")

# 4. MAQUETADO DE LA WEB SEGMENTADO POR TIPO DE PRODUCTO
html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Panel de Emisores por Tipo de Producto</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-950 text-gray-100 min-h-screen font-sans">
    <div class="max-w-7xl mx-auto px-4 py-8">
        
        <header class="border-b border-gray-800 pb-6 mb-12 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
            <div>
                <h1 class="text-3xl font-extrabold bg-gradient-to-r from-blue-400 via-indigo-400 to-emerald-400 bg-clip-text text-transparent">
                    Monitoreo de Emisores Segmentado
                </h1>
                <p class="text-gray-400 text-sm mt-1">Clasificación por Renta Fija, Variable, Fondos y Alternativos</p>
            </div>
            <div class="bg-gray-900 border border-gray-800 px-4 py-2 rounded-xl text-xs font-mono text-gray-400">
                Estrategia de Portafolio Activa
            </div>
        </header>
"""

# Generar bloques horizontales respetando el orden estricto
for producto in orden_productos:
    emisores_del_producto = datos_centralizados[producto]
    
    # Solo pintamos la sección de producto si contiene emisores activos con alertas/noticias
    hay_info = any(em["hechos"] or em["noticias"] for em in emisores_del_producto.values())
    if not hay_info:
        continue

    # Color diferenciador para cada tipo de activo
    badge_color = "text-blue-400 border-blue-500/30 bg-blue-500/5"
    if producto == "Renta Variable": badge_color = "text-purple-400 border-purple-500/30 bg-purple-500/5"
    elif producto == "Fondos": badge_color = "text-emerald-400 border-emerald-500/30 bg-emerald-500/5"
    elif producto == "Alternativos": badge_color = "text-amber-400 border-amber-500/30 bg-amber-500/5"

    html_content += f"""
        <section class="mb-12">
            <div class="flex items-center gap-3 mb-6 border-b border-gray-800 pb-2">
                <h2 class="text-xl font-bold tracking-tight text-white">{producto}</h2>
                <span class="border text-xs px-2.5 py-0.5 rounded-md font-medium {badge_color}">Clasificación Activa</span>
            </div>
            
            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
    """

    for emisor, contenido in emisores_del_producto.items():
        hechos = contenido["hechos"]
        noticias = contenido["noticias"]
        
        if not hechos and not noticias: continue

        html_content += f"""
                <div class="bg-gray-900 rounded-xl border border-gray-800 p-5 flex flex-col justify-between shadow-lg">
                    <div>
                        <div class="pb-2 mb-3 border-b border-gray-800/60 flex justify-between items-center">
                            <h3 class="text-md font-bold text-gray-100 tracking-wide uppercase">{html.escape(emisor)}</h3>
                        </div>
        """

        if hechos:
            html_content += """
                        <div class="mb-4">
                            <h4 class="text-[11px] font-bold text-amber-500 uppercase tracking-wider mb-2">🚨 Hechos SMV</h4>
                            <div class="space-y-2">"""
            for h in hechos:
                html_content += f"""
                                <div class="bg-amber-500/5 border-l-2 border-amber-500 p-2 rounded-r-md text-xs">
                                    <a href="{h['link']}" target="_blank" rel="noopener noreferrer" class="font-medium text-gray-300 hover:text-amber-400 line-clamp-2">{html.escape(h['titulo'])}</a>
                                    <span class="text-[10px] text-gray-500 block mt-0.5 font-mono">{h['fecha']}</span>
                                </div>"""
            html_content += """</div></div>"""

        if noticias:
            html_content += """
                        <div>
                            <h4 class="text-[11px] font-bold text-emerald-400 uppercase tracking-wider mb-2">📰 Prensa</h4>
                            <div class="space-y-2">"""
            for n in noticias:
                html_content += f"""
                                <div class="bg-emerald-500/5 border-l-2 border-emerald-500 p-2 rounded-r-md text-xs">
                                    <a href="{n['link']}" target="_blank" rel="noopener noreferrer" class="font-medium text-gray-300 hover:text-emerald-400 line-clamp-2">{html.escape(n['titulo'])}</a>
                                    <div class="flex justify-between text-[10px] text-gray-500 mt-0.5 font-mono">
                                        <span>{html.escape(n['fuente'])}</span>
                                        <span>{n['fecha']}</span>
                                    </div>
                                </div>"""
            html_content += """</div></div>"""

        html_content += """
                    </div>
                    <div class="text-[10px] text-gray-600 text-right mt-4 font-mono">Consolidado</div>
                </div>
        """

    html_content += """
            </div>
        </section>
    """

html_content += """
        <footer class="mt-16 pt-6 border-t border-gray-900 text-center text-xs text-gray-600">
            Filtro regulatorio segmentado por categorías de inversión.
        </footer>
    </div>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("¡index.html estructurado por categorías con éxito!")
