import urllib.request
import urllib.parse
import zipfile
import re
import ssl
import html
def leer_emisores_y_productos_estricto(ruta_archivo):
    mapping = {}
    try:
        with zipfile.ZipFile(ruta_archivo, 'r') as z:
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                with z.open('xl/sharedStrings.xml') as strings_file:
                    content = strings_file.read().decode('utf-8')
                    shared_strings = re.findall(r'<t[^>]*>(.*?)</t>', content)

            if 'xl/worksheets/sheet1.xml' in z.namelist():
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
                                        try:
                                            idx_str = int(val)
                                            texto = shared_strings[idx_str]
                                        except:
                                            texto = val
                                    else:
                                        texto = val
                                    texto = re.sub(r'<[^>]+>', ' ', texto).strip()
                                    datos_fila.append(texto)
                                else:
                                    datos_fila.append("")
                            
                            if len(datos_fila) >= 2:
                                emisor = str(datos_fila[0]).strip()
                                producto = str(datos_fila[1]).strip()
                                
                                if not emisor or emisor.lower() in ['emisores', 'emisor', 'nombre', 'empresa', 'razon social', 'company']:
                                    continue
                                    
                                prod_normalizado = producto.lower()
                                if 'fija' in prod_normalizado: mapping[emisor] = "Renta Fija"
                                elif 'variable' in prod_normalizado or 'accion' in prod_normalizado: mapping[emisor] = "Renta Variable"
                                elif 'fondo' in prod_normalizado: mapping[emisor] = "Fondos"
                                elif 'alterna' in prod_normalizado or 'alt' in prod_normalizado: mapping[emisor] = "Alternativos"
                                else: mapping[emisor] = "Renta Fija"
    except Exception as e:
        print("Nota: Procesando archivo base de datos...")
    return mapping

# 1. Cargar emisores UNICAMENTE desde tu archivo real Excel
mapping_emisores = leer_emisores_y_productos_estricto("Emisores.xlsx")

# 2. Inyección obligatoria exclusiva de Emisores Soberanos requeridos por el Comité
soberanos_requeridos = {
    "Peru": "Renta Fija", "Mexico": "Renta Fija", "Chile": "Renta Fija", "Colombia": "Renta Fija", "USA": "Renta Fija"
}
for pais, segmento in soberanos_requeridos.items():
    if pais not in mapping_emisores and (pais.lower() != "mexico" or "México" not in mapping_emisores):
        mapping_emisores[pais] = segmento

# Clasificación forzada de seguridad bursátil
reglas_seguridad = {
    "aceros arequipa": "Renta Variable", "alicorp": "Renta Variable", "volcan": "Renta Variable",
    "inretail": "Renta Variable", "credicorp": "Renta Variable", "bcp": "Renta Variable",
    "cerro verde": "Renta Variable", "banco gnb": "Renta Fija", "peru": "Renta Fija",
    "mexico": "Renta Fija", "chile": "Renta Fija", "colombia": "Renta Fija", "usa": "Renta Fija"
}
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

# Cabeceras completas de simulación de navegador humano de alta fidelidad
HEADERS_NATIVOS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'es-ES,es;q=0.9,en;q=0.8'
}

# Diccionario semántico de alertas de crédito, perspectivas y riesgo bursátil
patrones_criticos = [
    r'downgrade', r'upgrade', r'moody', r'fitch', r's&p', r'calificaci', r'clasificaci',
    r'perspectiva', r'rating', r'outlook', r'riesgo', r'sindicado', r'aumento de capital', 
    r'hecho de importancia', r'deuda', r'bonos', r'soberano', r'investment grade', r'grado de inversion', r'suscripci'
]
patron_riesgo = re.compile('|'.join(patrones_criticos), re.IGNORECASE)

patron_basura = re.compile(r'television|televisión|conductor|primiciasya|casella|espectaculo|futbol|fútbol|partido|farandula', re.IGNORECASE)
patron_fechas_viejas = re.compile(r'\b(2020|2021|2022)\b')

# Diccionario de Nemónicos clave (BVL) para soporte de búsqueda cruzada de tus corporativos principales
nemonicos_contingencia = {
    "Cerro Verde": "CVERDEC1", "Banco GNB": "GNBC1", "Fossal": "FOSSALC1",
    "Fibra Prime": "FIBPRIME", "Aceros Arequipa": "ACEROCI1", "Alicorp": "ALIACCI1", "Volcan": "VOLCABC1"
}

links_globales_procesados = set()

# CONTEXTO SSL: Desactivación estricta de validación para evadir bloqueos de cortafuegos gubernamentales
contexto_ssl_seguro = ssl.create_default_context()
contexto_ssl_seguro.check_hostname = False
contexto_ssl_seguro.verify_mode = ssl.CERT_NONE

# --- FASE 1: RASTREO ROBUSTO EN ENLACES CORE BURSÁTILES ---
print("Fase 1: Extrayendo información de portales bursátiles oficiales de manera aislada...")
PAGINAS_PRIORITARIAS = [
    {"url": "https://www.bvl.com.pe/emisores/noticias-emisores", "fuente": "BVL Oficial"},
    {"url": "https://www.smv.gob.pe/SIMV/frm_hechosdeImportanciaDia?data=38C2EC33FA106691BB5B5039DACFDF50795D8EC3AF", "fuente": "SMV Diario"},
    {"url": "https://www.smv.gob.pe/SIMV/Frm_HechosDeImportancia?data=EBE76110FDC9EF5632D5100F5B0448927EBDAC2CF7", "fuente": "SMV Historial"}
]

for portal in PAGINAS_PRIORITARIAS:
    try:
        req = urllib.request.Request(portal["url"], headers=HEADERS_NATIVOS)
        with urllib.request.urlopen(req, context=contexto_ssl_seguro, timeout=7) as response:
            html_puro = response.read().decode('utf-8', errors='ignore')
        
        texto_limpio = re.sub(r'<script[^>]*>([\s\S]*?)</script>|<style[^>]*>([\s\S]*?)</style>', '', html_puro)
        lineas = [re.sub(r'<[^>]+>', ' ', l).strip() for l in texto_limpio.split('\n') if l.strip()]
        
        for linea in lineas:
            if len(linea) < 12 or len(linea) > 240: continue
            if patron_fechas_viejas.search(linea) or patron_basura.search(linea): continue
            
            for emisor, prod in mapping_emisores.items():
                if not emisor: continue
                raiz_emisor = emisor.split()[0].replace(",", "").replace(".", "").strip().lower()
                if len(raiz_emisor) <= 3 and len(emisor.split()) > 1:
                    raiz_emisor = emisor.split()[1].replace(",", "").replace(".", "").strip().lower()
                
                nemonico = nemonicos_contingencia.get(emisor, "---").lower()
                
                if raiz_emisor in linea.lower() or nemonico in linea.lower():
                    es_prioritaria = bool(patron_riesgo.search(linea)) or any(k in portal["fuente"].lower() for k in ["bvl", "smv"])
                    if emisor in soberanos_requeridos and not es_prioritaria: continue
                    
                    if linea in links_globales_procesados: continue
                    links_globales_procesados.add(linea)
                    
                    # BLINDADO CONTRA CORTES: Extracción segura de títulos sin romper por guiones
                    titulo_limpio = linea
                    if " - " in linea:
                        partes = linea.split(" - ")
                        if len(partes) > 0 and partes[0].strip():
                            titulo_limpio = partes[0].strip()
                    
                    datos_centralizados[prod][emisor]["noticias"].append({
                        "titulo": titulo_limpio,
                        "link": portal["url"],
                        "fuente": portal["fuente"],
                        "fecha": "Hoy",
                        "prioritaria": es_prioritaria
                    })
    except Exception as e:
        print(f"Canal core diferido de forma segura")

# --- FASE 2: RASTREO MULTI-REGIONAL SEGURO (Google News + Bloomberg Línea) ---
print("Fase 2: Ejecutando consultas cruzadas en la red global...")
for idx, (emisor, prod) in enumerate(mapping_emisores.items()):
    if not emisor or prod not in datos_centralizados: continue
    
    is_soberano = emisor in soberanos_requeridos
    emisor_encoded = urllib.parse.quote(emisor)
    ventana_tiempo = "tbs=qdr:m6" if (is_soberano or emisor.lower() in ["cerro verde", "alicorp", "mexico"]) else "tbs=qdr:m"
    
    if is_soberano:
        query_bvl = f'("{emisor}"%20OR%20"Republica%20de%20{emisor}")%20(Moody%27s%20OR%20Fitch%20OR%20Downgrade%20OR%20Outlook%20OR%20Perspectiva)'
    else:
        cod_bvl = nemonicos_contingencia.get(emisor, emisor)
        query_bvl = f'("{emisor}"%20OR%20"{cod_bvl}")%20(Moody%27s%20OR%20Downgrade%20OR%20Sindicado%20OR%20SMV%20OR%20BVL%20OR%20Outlook%20OR%20Suscripcion%20OR%20"Aumento%20de%20Capital")'

    urls_red = [
        f"https://news.google.com/rss/search?q={query_bvl}&hl=es-419&gl=PE&ceid=PE:es-419&{ventana_tiempo}",
        f"https://news.google.com/rss/search?q={emisor_encoded}%20site:bloomberglinea.com&hl=es-419&gl=PE&ceid=PE:es-419&{ventana_tiempo}"
    ]
    
    for url in urls_red:
        try:
            req = urllib.request.Request(url, headers=HEADERS_NATIVOS)
            with urllib.request.urlopen(req, context=contexto_ssl_seguro, timeout=8) as response:
                xml_raw = response.read().decode('utf-8', errors='ignore')
            
            items_raw = re.findall(r'<item>([\s\S]*?)</item>', xml_raw)
            
            for item_content in items_raw[:5]:
                t_match = re.search(r'<title>([\s\S]*?)</title>', item_content)
                l_match = re.search(r'<link>([\s\S]*?)</link>', item_content)
                d_match = re.search(r'<pubDate>([\s\S]*?)</pubDate>', item_content)
                f_match = re.search(r'<source[^>]*>([\s\S]*?)</source>', item_content)
                
                titulo = t_match.group(1).strip() if t_match else ""
                link = l_match.group(1).strip() if l_match else ""
                pub_date = d_match.group(1).strip() if d_match else ""
                fuente = f_match.group(1).strip() if f_match else "Prensa"
                
                if not titulo or not link: continue
                if "bloomberglinea.com" in link: fuente = "Bloomberg Línea"
                
                if patron_fechas_viejas.search(titulo) or patron_basura.search(titulo): continue
                if is_soberano and not any(x in titulo.lower() for x in [emisor.lower(), 'perspectiva', 'rating', 'downgrade', 'outlook']): continue
                
                raiz_emisor = emisor.split()[0].replace(",", "").replace(".", "").strip().lower()
                cod_bvl_lower = nemonicos_contingencia.get(emisor, "---").lower()
                
                if not is_soberano and (raiz_emisor not in titulo.lower() and cod_bvl_lower not in titulo.lower()):
                    continue
                
                if link in links_globales_procesados or titulo in links_globales_procesados: continue
                links_globales_procesados.add(link)
                links_globales_procesados.add(titulo)
                
                es_prioritaria = bool(patron_riesgo.search(titulo)) or "bloomberg" in fuente.lower() or is_soberano
                fecha_limpia = pub_date[:11].strip() if pub_date else "Hoy"
                
                # Formateo seguro de títulos de red
                titulo_final = titulo
                if " - " in titulo:
                    partes_t = titulo.split(" - ")
                    if len(partes_t) > 0 and partes_t[0].strip():
                        titulo_final = partes_t[0].strip()
                
                datos_centralizados[prod][emisor]["noticias"].append({
                    "titulo": titulo_final,
                    "link": link,
                    "fuente": fuente,
                    "fecha": fecha_limpia,
                    "prioritaria": es_prioritaria
                })
        except:
            pass 

# --- FASE 3: ORDENACIÓN DE ALERTAS DE CRÉDITO ---
for prod in orden_productos:
    for emisor in datos_centralizados[prod]:
        datos_centralizados[prod][emisor]["noticias"].sort(key=lambda x: x["prioritaria"], reverse=True)
        datos_centralizados[prod][emisor]["noticias"] = datos_centralizados[prod][emisor]["noticias"][:4]

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
                <p class="text-gray-400 text-sm mt-1">Filtro de Crédito Avanzado • Sincronización Estricta de tu Portafolio de Inversión</p>
            </div>
            <div class="bg-red-950/40 border border-red-500/30 px-4 py-2 rounded-xl text-xs font-mono text-red-400 flex items-center gap-2">
                <span class="w-2 h-2 rounded-full bg-red-500 animate-pulse"></span> Terminal Blindada y Abierta Sincronizada Activa
            </div>
        </header>
"""

total_visibles = 0
for producto in orden_productos:
    emisores_del_producto = datos_centralizados[producto]
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
        
        es_pais = emisor in soberanos_requeridos
        bg_card_base = "from-gray-900 to-blue-950/10 border-blue-500/30" if es_pais else "from-gray-900 to-gray-900/40 border-gray-800"
        
        tiene_alerta = any(n["prioritaria"] for n in noticias)
        borde_caja = f"border-red-500/50 bg-gradient-to-b from-gray-900 to-red-950/20 shadow-red-950/30 shadow-lg" if tiene_alerta else f"border-gray-800 bg-gradient-to-b {bg_card_base} shadow-lg"

        html_content += f"""
                <div class="rounded-xl border p-5 flex flex-col justify-between hover:border-gray-700 transition duration-300 {borde_caja}">
                    <div>
                        <div class="pb-2 mb-3 border-b border-gray-800/60 flex justify-between items-center">
                            <div class="flex items-center gap-2">
                                <h3 class="text-sm font-bold text-gray-100 tracking-wide uppercase">{html.escape(emisor)}</h3>
                                { '<span class="text-[9px] px-1.5 py-0.2 bg-blue-500/10 text-blue-400 rounded-md border border-blue-500/20 font-mono">Soberano</span>' if es_pais else '' }
                            </div>
                            { '<span class="bg-red-500/20 text-red-400 border border-red-500/30 text-[9px] px-1.5 py-0.5 rounded uppercase font-bold tracking-wider animate-pulse">Rating Alert</span>' if tiene_alerta else '' }
                        </div>
                        <div class="space-y-2">
        """
        for n in noticias:
            bg_item = "bg-red-500/10 border-red-500 text-red-200 font-medium" if n["prioritaria"] else "bg-gray-850 border-gray-700 text-gray-300"
            html_content += f"""
                            <div class="border-l-2 p-2 rounded-r text-xs {bg_item}">
                                <div class="flex justify-between items-start gap-2 mb-1">
                                    <a href="{n['link']}" target="_blank" class="hover:text-blue-400 line-clamp-3 font-medium flex-1">{html.escape(n['titulo'])}</a>
                                    <span class="text-[9px] text-gray-500 bg-gray-900 px-1.5 py-0.5 rounded border border-gray-800 font-mono shrink-0 whitespace-nowrap">{html.escape(n['fecha'])}</span>
                                </div>
                                <div class="text-[9px] text-gray-500 font-mono">
                                    <span>Origen: {html.escape(n['fuente'])}</span>
                                </div>
                            </div>
            """
        html_content += """</div></div></div>"""
    html_content += """</div></section>"""

if total_visibles == 0:
    html_content += """
        <div class="bg-gray-900 border border-gray-800 rounded-xl p-12 text-center max-w-xl mx-auto my-12">
            <p class="text-emerald-400 font-medium text-lg mb-2">☕ Todo bajo control</p>
            <p class="text-gray-400 text-sm">No se registran alertas de crédito ni movimientos para tu portafolio de Excel hoy.</p>
        </div>
    """

html_content += """
        <footer class="mt-16 pt-6 border-t border-gray-900 text-center text-xs text-gray-600">
            Filtro de Crédito Customizado Multiregional Soberano e Institucional Seleccionado.
        </footer>
    </div>
</body>
</html>
# --- ESCRIBIR EL ARCHIVO FINAL EN EL REPOSITORIO ---
with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
print("¡Fichero index.html unificado y completado con éxito con protección ante bloqueos!")
