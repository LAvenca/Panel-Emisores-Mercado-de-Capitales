import urllib.request
import xml.etree.ElementTree as ET
import html

# 1. Leer los emisores desde el archivo
with open("emisores.txt", "r", encoding="utf-8") as f:
    emisores = [line.strip() for line in f if line.strip()]

noticias_totales = []

# 2. Buscar noticias en el RSS de Google News para cada emisor
for emisor in emisores:
    # Codificar el nombre para la URL
    query = urllib.parse.quote(emisor)
    url = f"https://news.google.com/rss/search?q={query}&hl=es-419&gl=US&ceid=US:es-419"
    
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            xml_data = response.read()
            
        root = ET.fromstring(xml_data)
        # Tomar solo las 3 noticias más recientes de cada emisor
        count = 0
        for item in root.findall('.//item'):
            if count >= 3:
                break
            titulo = item.find('title').text
            link = item.find('link').text
            fecha = item.find('pubDate').text
            fuente = item.find('source').text if item.find('source') is not None else "Fuente externa"
            
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

# 3. Generar el diseño HTML dinámico (Moderno y Limpio)
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
                <h1 class="text-3xl font-bold bg-gradient-to-r from-green-400 to-blue-500 bg-clip-text text-transparent">
                    Monitoreo de Emisores
                </h1>
                <p class="text-gray-400 text-sm mt-1">Actualizado automáticamente vía GitHub Actions</p>
            </div>
            <div class="bg-gray-800 px-4 py-2 rounded-lg text-xs text-gray-400 border border-gray-700">
                Última actualización: <span class="font-mono text-green-400">Diaria automática</span>
            </div>
        </header>

        <main class="grid gap-6 md:grid-cols-2">
"""

# Inyectar las noticias al HTML
for n in noticias_totales:
    html_content += f"""
            <div class="bg-gray-800 p-6 rounded-xl border border-gray-700 hover:border-green-500 transition duration-300 flex flex-col justify-between">
                <div>
                    <div class="flex justify-between items-center mb-3">
                        <span class="bg-green-500/10 text-green-400 text-xs font-semibold px-2.5 py-1 rounded">
                            {html.escape(n['emisor'])}
                        </span>
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
            </div>
    """

html_content += """
        </main>
        <footer class="mt-16 pt-6 border-t border-gray-800 text-center text-xs text-gray-500">
            Sistema de automatización de portafolio.
        </footer>
    </div>
</body>
</html>
"""

# 4. Guardar el archivo index.html
with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("¡index.html generado con éxito!")
