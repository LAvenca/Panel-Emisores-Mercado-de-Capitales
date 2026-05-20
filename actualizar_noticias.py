import urllib.request
import xml.etree.ElementTree as ET
import html
import zipfile

def leer_emisores_excel(ruta_archivo):
    emisores = []
    try:
        with zipfile.ZipFile(ruta_archivo, 'r') as z:
            with z.open('xl/worksheets/sheet1.xml') as sheet:
                tree = ET.parse(sheet)
                root = tree.getroot()
                namespaces = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                
                shared_strings = []
                if 'xl/sharedStrings.xml' in z.namelist():
                    with z.open('xl/sharedStrings.xml') as strings_file:
                        str_tree = ET.parse(strings_file)
                        str_root = str_tree.getroot()
                        shared_strings = [el.text for el in str_root.findall('.//ns:t', namespaces)]

                for row in root.findall('.//ns:row', namespaces):
                    cell = row.find('ns:c', namespaces)
                    if cell is not None:
                        val_el = cell.find('ns:v', namespaces)
                        if val_el is not None:
                            val = val_el.text
                            if cell.get('t') == 's' and shared_strings:
                                idx = int(val)
                                emisor_name = shared_strings[idx]
                            else:
                                emisor_name = val
                            
                            if emisor_name and emisor_name.strip():
                                emisores.append(emisor_name.strip())
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
    
    if emisores and emisores[0].lower() in ['emisores', 'emisor', 'nombre', 'empresa', 'empresas']:
        emisores.pop(0)
        
    return list(set(emisores))

emisores = leer_emisores_excel("emisores.xlsx")
noticias_totales = []

for emisor in emisores:
    query = urllib.parse.quote(emisor)
    url = f"https://news.google.com/rss/search?q={query}&hl=es-419&gl=US&ceid=US:es-419"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
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
        print(f"Error con {emisor}: {e}")

html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Noticias de mi Portafolio</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-900 text-gray-100 min-h-screen font-sans">
    <div class="max-w-5xl mx-auto px-4 py-8">
        <header class="border-b border-gray-800 pb-6 mb-8 flex justify-between items-center">
            <div>
                <h1 class="text-3xl font-bold bg-gradient-to-r from-green-400 to-blue-500 bg-clip-text text-transparent">Panel de Emisores</h1>
                <p class="text-gray-400 text-sm mt-1">Actualizado automáticamente desde Excel</p>
            </div>
        </header>
        <main class="grid gap-6 md:grid-cols-2">"""

for n in noticias_totales:
    html_content += f"""
            <div class="bg-gray-800 p-6 rounded-xl border border-gray-700 hover:border-green-500 transition duration-300 flex flex-col justify-between">
                <div>
                    <div class="flex justify-between items-center mb-3">
                        <span class="bg-green-500/10 text-green-400 text-xs font-semibold px-2.5 py-1 rounded">{html.escape(n['emisor'])}</span>
                        <span class="text-gray-500 text-xs">{html.escape(n['fuente'])}</span>
                    </div>
                    <h2 class="text-lg font-semibold text-white mb-2 hover:text-green-400">
                        <a href="{n['link']}" target="_blank" rel="noopener noreferrer">{html.escape(n['titulo'])}</a>
                    </h2>
                </div>
                <div class="mt-4 pt-3 border-t border-gray-700/50 flex justify-between items-center text-xs text-gray-400">
                    <span>{n['fecha'][:16]}</span>
                    <a href="{n['link']}" target="_blank" class="text-blue-400 hover:underline">Leer más →</a>
                </div>
            </div>"""

html_content += """</main></div></body></html>"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
print("¡index.html generado!")
