import urllib.request
import urllib.parse
import zipfile
import re
import ssl
import html

# --- CONFIGURACIÓN MAESTRA ---
HEADERS_NATIVOS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
}

# Fuentes obligatorias de monitoreo
PAGINAS_PRIORITARIAS = [
    {"url": "https://www.bvl.com.pe/emisores/noticias-emisores", "fuente": "BVL Oficial"},
    {"url": "https://www.smv.gob.pe/SIMV/frm_hechosdeImportanciaDia?data=38C2EC33FA106691BB5B5039DACFDF50795D8EC3AF", "fuente": "SMV Diario"},
    {"url": "https://www.smv.gob.pe/SIMV/Frm_HechosDeImportancia?data=EBE76110FDC9EF5632D5100F5B0448927EBDAC2CF7", "fuente": "SMV Historial"},
    {"url": "https://www.moodys.com/entity/489500/overview", "fuente": "Moody's Radar"}
]

# Diccionario de alertas críticas
patrones_criticos = [
    r'downgrade', r'upgrade', r'moody', r'fitch', r's&p', r'calificaci', r'clasificaci',
    r'perspectiva', r'rating', r'riesgo', r'sindicado', r'aumento de capital', 
    r'hecho de importancia', r'deuda', r'bonos', r'soberano', r'investment grade', r'grado de inversion', r'credito'
]
patron_riesgo = re.compile('|'.join(patrones_criticos), re.IGNORECASE)

# --- LÓGICA DE EXTRACCIÓN (Resto del código) ---
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
                                        try: texto = shared_strings[int(val)]
                                        except: texto = val
                                    else: texto = val
                                    texto = re.sub(r'<[^>]+>', ' ', texto).strip()
                                    datos_fila.append(texto)
                                else: datos_fila.append("")
                            if len(datos_fila) >= 2:
                                emisor = str(datos_fila[0]).strip()
                                producto = str(datos_fila[1]).strip()
                                if emisor and emisor.lower() not in ['emisores', 'emisor', 'nombre', 'empresa']:
                                    mapping[emisor] = "Renta Variable" if 'variable' in producto.lower() else "Renta Fija"
    except: pass
    return mapping

mapping_emisores = leer_emisores_y_productos_estricto("Emisores.xlsx")
# (Añadir resto de lógica de procesamiento y maquetado HTML aquí...)
