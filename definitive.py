import os
import re
import pandas as pd
from openpyxl import Workbook
from PyPDF2 import PdfReader
from datetime import datetime

directorio = r'C:\Users\user\Documents\Digitalizados_2021\11001333501020180005000\01PrimeraInstancia\C01Principal1'

archivos_info = []

def limpiar_nombre_archivo(nombre_archivo):
    nombre_sin_extension = os.path.splitext(nombre_archivo)[0]
    nombre_limpio = re.sub(r'\d+', '', nombre_sin_extension)
    nombre_limpio = nombre_limpio.replace('_', ' ').strip()
    return nombre_limpio

for archivo in os.listdir(directorio):
    ruta_archivo = os.path.join(directorio, archivo)
    if os.path.isfile(ruta_archivo):
        nombre_archivo_original = archivo
        nombre_archivo_limpio = limpiar_nombre_archivo(nombre_archivo_original)
        fecha_creacion = datetime.fromtimestamp(os.path.getctime(ruta_archivo)).strftime('%d/%m/%Y')
        size_file = os.path.getsize(ruta_archivo) / 1024
        size_file = f"{size_file:.0f} KB"
        
        num_paginas = 'N/A'
        if archivo.lower().endswith('.pdf'):
            try:
                reader = PdfReader(ruta_archivo)
                num_paginas = len(reader.pages)
            except:
                num_paginas = 'Error al leer'

        archivos_info.append({
            'Nombre del archivo': nombre_archivo_limpio,
            'Fecha de creación': fecha_creacion,
            'Número de páginas': num_paginas,
            'Tamaño del archivo': size_file
        })

df = pd.DataFrame(archivos_info)

ruta_excel = os.path.join(directorio, '00IndiceElectronico.xlsx')
df.to_excel(ruta_excel, index=False)

print(f'Información de archivos guardada en {ruta_excel}')