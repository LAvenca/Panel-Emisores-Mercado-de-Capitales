import pandas as pd
import re
import html
import ssl
import urllib.request

# --- CONFIGURACIÓN ---
patrones_criticos = ['downgrade', 'upgrade', 'moody', 'fitch', 'rating', 'riesgo', 'deuda', 'credito']
PAGINAS_PRIORITARIAS = [
    {"url": "https://www.bvl.com.pe/emisores/noticias-emisores", "fuente": "BVL Oficial"},
    {"url": "https://www.smv.gob.pe/SIMV/frm_hechosdeImportanciaDia?data=38C2EC33FA106691BB5B5039DACFDF50795D8EC3AF", "fuente": "SMV Diario"}
]

# --- LECTURA SEGURA ---
def cargar_emisores():
    try:
        # Esto lee el Excel directamente usando la columna correcta
        df = pd.read_excel("Emisores.xlsx")
        # Asegúrate de que tu Excel tenga cabeceras 'Emisor' y 'Segmento'
        return dict(zip(df['Emisor'], df['Segmento']))
    except Exception as e:
        print(f"ERROR AL LEER EXCEL: {e}")
        return {"PERU": "Renta Fija"} # Valor de respaldo

emisores_dict = cargar_emisores()

# --- GENERACIÓN DE HTML ---
html_content = ["<html><body><h1>Monitoreo</h1>"]
for emisor, seg in emisores_dict.items():
    html_content.append(f"<div>{seg} - {emisor}</div>")
html_content.append("</body></html>")

with open("index.html", "w", encoding="utf-8") as f:
    f.write("\n".join(html_content))

print("Proceso finalizado correctamente.")
