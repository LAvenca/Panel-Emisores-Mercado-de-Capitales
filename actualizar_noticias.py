import urllib.request
import xml.etree.ElementTree as ET
import html
import zipfile
import re

def leer_emisores_y_productos_robusto(ruta_archivo):
    mapping = {}
    try:
        with zipfile.ZipFile(ruta_archivo, 'r') as z:
            # 1. Extraer todos los textos indexados del Excel
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                with z.open('xl/sharedStrings.xml') as strings_file:
                    content = strings_file.read().decode('utf-8')
                    shared_strings = re.findall(r'<t[^>]*>(.*?)</t>', content)

            # 2. Leer la estructura de las celdas por fila
            with z.open('xl/worksheets/sheet1.xml') as sheet:
                content = sheet.read().decode('utf-8')
                
                # Expresión regular para capturar cada fila completa
                filas = re.findall(r'<row[^>]*>(.*?)</row>', content)
                
                for idx, fila in enumerate(filas):
                    # Capturar el contenido de cada celda individualmente dentro de la fila
                    celdas = re.findall(r'<c[^>]*>(.*?)</c>|<c[^>]*/>', fila)
                    
                    # Necesitamos al menos dos celdas con datos en la fila
                    if len(celdas) >= 2:
                        datos_fila = []
                        for celda in celdas[:2]:
                            v_match = re.search(r'<v>(.*?)</v>', celda)
                            if v_match:
                                val = v_match.group(1)
                                # Si la celda indica que es tipo 's' (shared string)
                                if 't="s"' in celda or (celda == celdas[0] and 't="s"' in filas[idx]):
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
                            
                            # Omitir la primera fila si es el encabezado de las columnas
                            if idx == 0 and emisor.lower() in ['emisores', 'emisor', 'nombre', 'empresa']:
                                continue
                                
                            if emisor and producto:
                                # Normalizar el nombre del producto para evitar errores de escritura en el Excel
                                prod_normalizado = producto.strip().lower()
                                if 'fija' in prod_normalizado: mapping[emisor] = "Renta Fija"
                                elif 'variable' in prod_normalizado: mapping[emisor] = "Renta Variable"
                                elif 'fondo' in prod_normalizado: mapping[emisor] = "Fondos"
                                elif 'alterna' in prod_normalizado or 'alt' in prod_normalizado: mapping[emisor] = "Alternativos"
                                else: mapping[emisor] = "Renta Fija" # Categoría por defecto
    except Exception as e:
        print(f"Error procesando el Excel: {e}")
    return mapping

# Cargar la configuración desde el archivo Excel subido
mapping_emisores = leer_emisores_y_productos_robusto("Emisores.xlsx")

# Plan de respaldo si la estructura del archivo sigue resistiéndose a la lectura nativa
if not mapping_emisores:
    print("Aviso: No se pudo mapear el Excel. Activando lista de emisores de respaldo...")
    mapping_emisores = {
        "Pacífico Seguros": "Renta Fija",
        "Rímac Seguros": "Renta Fija",
        "Interseguro": "Renta Variable",
        "Credicorp": "Renta Variable",
        "Intercorp": "Fondos",
        "Fibra Prime": "Alternativos"
    }

print(f"Estructura final de emisores cargada: {mapping_emisores}")

orden_productos = ["Renta Fija", "Renta Variable", "Fondos", "Alternativos"]
datos_centralizados = {prod: {} for prod in orden_productos}

for emisor, prod in mapping_emisores.items():
    if prod in datos_centralizados:
        datos_centralizados[prod][emisor] = {"hechos": [], "noticias": []}

# 3. Descargar Hechos Relevantes desde la SMV
print("Buscando información regulatoria en la SMV...")
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
    print(f"Error SMV: {e}")

# 4. Descargar noticias desde Google News de manera global
print("Buscando información en prensa financiera...")
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
        print(f"Error noticias: {e}")

# 5. CONSTRUCCIÓN DE LA WEB DINÁMICA
html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Monitoreo de Emisores Segmentado</title>
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

for producto in orden_productos:
    emisores_del_producto = datos_centralizados[producto]
    
    # Mostrar la sección si tiene emisores configurados para que la grilla nunca se quede vacía
    if not emisores_del_producto:
        continue

    badge_color = "text-blue-400 border-blue-500/30 bg-blue-500/5"
    if producto == "Renta Variable": badge_color = "text-purple-400 border-purple-500/30 bg-purple-500/5"
    elif producto == "Fondos": badge_color = "text-emerald-400 border-emerald-500/30 bg-emerald-500/5"
    elif producto == "Alternativos": badge_color = "text-amber-400 border-amber-500/30 bg-amber-500/5"

    html_content += f"""
        <section class="mb-12">
            <div class="flex items-center gap-3 mb-6 border-b border-gray-800 pb-2">
                <h2 class="text-xl font-bold text-white">{producto}</h2>
                <span class="border text-xs px-2.5 py-0.5 rounded-md font-medium {badge_color}">Clasificación Activa</span>
            </div>
            
            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
    """

    for emisor, contenido in list(emisores_del_producto.items()):
        hechos = contenido["hechos"]
        noticias = contenido["noticias"]
        
        # Si un emisor no tiene novedades hoy, le ponemos un aviso sutil dentro de su caja para mantener la grilla estructurada
        html_content += f"""
                <div class="bg-gray-900 rounded-xl border border-gray-800 p-5 flex flex-col justify-between shadow-lg hover:border-gray-700 transition">
                    <div>
                        <div class="pb-2 mb-3 border-b border-gray-800/60">
                            <h3 class="text-md font-bold text-gray-100 tracking-wide uppercase">{html.escape(emisor)}</h3>
                        </div>
        """

        if not hechos and not noticias:
            html_content += """
                        <p class="text-xs text-gray-600 italic py-4">Sin alertas de prensa ni hechos relevantes reportados en las últimas horas.</p>
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
