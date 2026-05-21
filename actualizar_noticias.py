import urllib.request
import xml.etree.ElementTree as ET
import html
import zipfile
import re

def leer_emisores_excel_robusto(ruta_archivo):
    emisores = []
    try:
        with zipfile.ZipFile(ruta_archivo, 'r') as z:
            # 1. Intentar cargar los strings compartidos (donde Excel suele guardar los textos)
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                with z.open('xl/sharedStrings.xml') as strings_file:
                    content = strings_file.read().decode('utf-8')
                    shared_strings = re.findall(r'<t[^>]*>(.*?)</t>', content)

            # 2. Leer la primera hoja
            with z.open('xl/worksheets/sheet1.xml') as sheet:
                content = sheet.read().decode('utf-8')
                # Buscar todas las filas
                filas = re.findall(r'<row[^>]*>(.*?)</row>', content)
                
                for fila in filas:
                    # Buscar la primera celda de la fila (Columna A)
                    celdas = re.findall(r'<c[^>]*>(.*?)</c>', fila)
                    if celdas:
                        primera_celda = celdas[0]
                        # Ver si es un string compartido o un valor directo
                        es_shared = 't="s"' in fila or 't="s"' in re.search(r'<c[^>]*>', fila).group(0) if re.search(r'<c[^>]*>', fila) else False
                        
                        v_match = re.search(r'<v>(.*?)</v>', primera_celda)
                        if v_match:
                            valor = v_match.group(1)
                            if es_shared and shared_strings:
                                try:
                                    texto = shared_strings[int(valor)]
                                except:
                                    texto = valor
                            else:
                                texto = valor
                            
                            # Limpiar etiquetas HTML residuales
                            texto_limpio = re.sub(r'<[^>]+>', '', texto).strip()
                            if texto_limpio:
                                emisores.append(texto_limpio)
    except Exception as e:
        print(f"Error al leer Excel: {e}")
    
    # Filtrar encabezados comunes si existen
    if emisores and emisores[0].lower() in ['emisores', 'emisor', 'nombre', 'empresa', 'empresas', 'columna1']:
        emisores.pop(0)
        
    return list(set([e for e in emisores if len(e) > 1]))

# Cargar emisores
emisores = leer_emisores_excel_robusto("Emisores.xlsx")
print(f"Emisores encontrados: {emisores}")

# Si por algún motivo el Excel falla, usar una lista de respaldo para que no quede vacío
if not emisores:
    emisores = ["Pacífico Seguros", "Interseguro", "Rímac Seguros"]
    print(f"Usando lista de respaldo por error en lectura de Excel: {emisores}")

noticias_totales = []

for emisor in emisores:
    query = urllib.parse.quote(emisor)
    url = f"https://news.google.com/rss/search?q={query}&hl=es-419&gl=US&ceid=US:es-419"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
        with urllib.request.urlopen(req) as response:
            xml_data = response.read()
        root = ET.fromstring(xml_data)
        count = 0
        for item in root.findall('.//item'):
            if count >= 3:
                break
            titulo = item.find('title').text
            link = item.find('link').text
            fecha = item.find('pubDate').text
            fuente = item.find('source').text if item.find('source') is not None else "Fuente"
            
            noticias_totales.append({
                "emisor": emisor,
                "titulo": titulo,
                "link": link,
                "fecha": fecha,
                "fuente": fuente
            })
            count += 1
    except Exception as e:
        print(f"Error buscando noticias para {emisor}: {e}")

# Diseño Web
html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Panel de Emisores</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-900 text-gray-100 min-h-screen font-sans">
    <div class="max-w-6xl mx-auto px-4 py-8">
        <header class="border-b border-gray-800 pb-6 mb-8">
            <h1 class="text-3xl font-bold bg-gradient-to-r from-green-400 to-blue-500 bg-clip-text text-transparent">Panel de Emisores</h1>
            <p class="text-gray-400 text-sm mt-1">Monitoreo del Mercado de Capitales — Actualizado automáticamente</p>
        </header>
        <main class="grid gap-6 md:grid-cols-2 lg:grid-cols-3">"""

if not noticias_totales:
    html_content += """
        <div class="col-span-full text-center py-12 text-gray-500">
            No se encontraron noticias recientes para los emisores listados.
        </div>"""

for n in noticias_totales:
    html_content += f"""
            <div class="bg-gray-800 p-5 rounded-xl border border-gray-700 hover:border-green-500 transition duration-300 flex flex-col justify-between shadow-lg">
                <div>
                    <div class="flex justify-between items-center mb-3">
                        <span class="bg-green-500/10 text-green-400 text-xs font-semibold px-2.5 py-1 rounded">
                            {html.escape(n['emisor'])}
                        </span>
                        <span class="text-gray-400 text-xs">{html.escape(n['fuente'])}</span>
                    </div>
                    <h2 class="text-base font-medium text-white mb-4 line-clamp-3 hover:text-green-400">
                        <a href="{n['link']}" target="_blank" rel="noopener noreferrer">{html.escape(n['titulo'])}</a>
                    </h2>
                </div>
                <div class="pt-3 border-t border-gray-700/50 flex justify-between items-center text-xs text-gray-400">
                    <span>{n['fecha'][:16]}</span>
                    <a href="{n['link']}" target="_blank" class="text-blue-400 font-medium hover:underline">Leer más →</a>
                </div>
            </div>"""

html_content += """</main></div></body></html>"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
print("¡index.html generado exitosamente!")
