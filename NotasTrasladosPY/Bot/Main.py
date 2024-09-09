import json
import os
import time
import datetime as dt
from os.path import join
import sys

# Obtén la ruta del directorio actual donde se encuentra el archivo
current_dir = os.path.dirname(os.path.abspath(__file__))

# Mueve un nivel hacia arriba
workfolder_path = os.path.dirname(current_dir)

archivo_excel = os.path.join(workfolder_path, 'Onedrive', 'reporte de notas.xls')

# Cambiar el directorio de trabajo al directorio del script
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Obtenemos la fecha actual en el formato deseado: dia-mes-año
today = dt.datetime.now().strftime("%d-%m-%Y")

# Ruta donde se guardarán los informes y el archivo ZIP
folder = join(workfolder_path, 'reports', today)

# Ruta completa del archivo ZIP
zip_path = join(folder, 'Resultado.zip')

if os.path.exists(archivo_excel):

    if not os.path.isfile(zip_path):

        # Ejecutar eprimer bloque
        exec(open('Configuracion.py').read())
        time.sleep(5)

        # Ejecutar el segundo bloque
        exec(open('CargaFiltrado.py', encoding='utf-8').read())
        time.sleep(5)

        # Ejecutar el tercer bloque
        exec(open('GenerarConsolidado.py').read())
        time.sleep(5)

        # Ejecutar el cuarto bloque
        exec(open('GeneracionNotas.py').read())
        time.sleep(5)

        # Ejecutar el quinto bloque
        exec(open('ConvertirNotasPDF.py').read())
        time.sleep(5)

        # Ejecutar el sexto bloque
        exec(open('ConsolidadoEpm.py').read())
        time.sleep(5)

        # Ejecutar el séptimo bloque
        exec(open('EpmPdf.py').read())
        time.sleep(5)

        # Ejecutar el octavo bloque
        exec(open('ConsolidadoComfama.py').read())
        time.sleep(5)

        # Ejecutar el noveno bloque
        exec(open('ComfamaPdf.py').read())
        time.sleep(5)

        # Ejecutar el décimo bloque
        exec(open('CrearZip.py').read())
        time.sleep(5)

        # Ejecutar el undécimo bloque
        exec(open('EnvioCorreo.py', encoding='utf-8').read())
        time.sleep(5)
    else:
        print("Ya existe el archivo zip")
        sys.exit(0)
else:
    exec(open('NoExiste.py', encoding='utf-8').read())