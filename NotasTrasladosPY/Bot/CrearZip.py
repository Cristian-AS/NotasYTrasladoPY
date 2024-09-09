from os import listdir
from os.path import join, isdir, relpath, exists, isfile
import datetime as dt
import zipfile
import os

# Obtén la ruta del directorio actual donde se encuentra el archivo
current_dir = os.path.dirname(os.path.abspath(__file__))

# Mueve un nivel hacia arriba
workfolder_path = os.path.dirname(current_dir)

# Obtenemos la fecha actual en el formato deseado: dia-mes-año
today = dt.datetime.now().strftime("%d-%m-%Y")

# Ruta donde se guardarán los informes y el archivo ZIP
folder = join(workfolder_path, 'reports', today)

# Ruta completa del archivo ZIP a crear
zip_path = join(folder, 'Resultado.zip')

# Lista de carpetas a incluir en la compresión
folders_to_compress = ['comfama', 'epm', 'notas']

# Verificar si el archivo ZIP ya existe
if not isfile(zip_path):
    # Si no existe, crear el archivo ZIP y agregar los archivos PDF a comprimir
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zfp:
        for folder_name in folders_to_compress:
            folder_path = join(folder, folder_name)
            if isdir(folder_path):
                for file in listdir(folder_path):
                    # Verificar si el archivo es un archivo PDF
                    if file.lower().endswith(".pdf"):
                        file_path = join(folder_path, file)
                        # Agregar el archivo PDF al archivo ZIP
                        zfp.write(filename=file_path, arcname=file)
    print("Zip file created.")
else:
    print("Zip file already exists.")