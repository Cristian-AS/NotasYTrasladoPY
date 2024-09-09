from os.path import join
from os import listdir
import os
import datetime as dt
import win32com.client
from pathlib import Path

# Obtén la ruta del directorio actual donde se encuentra el archivo
current_dir = os.path.dirname(os.path.abspath(__file__))

# Mueve un nivel hacia arriba
workfolder_path = os.path.dirname(current_dir)

# Variables
folder_name = 'comfama'

# Declaración de una variable para la fecha de hoy
today = dt.datetime.today().strftime("%d-%m-%Y")

# Ruta de la carpeta donde están almacenados los reportes a convertir
folder = join(workfolder_path, "reports", today, folder_name)

# Lista de archivos en la carpeta
files = [file.split('.')[0] for file in listdir(folder)]

# Procesamiento de cada archivo
for file in files:
    filepath = join(folder, f'{file}.xlsx')
    pdf_path = join(folder, f'{file}.pdf').replace('.xlsx', '')
    
    print(filepath)
    
    # Verifica si el archivo PDF ya existe
    if Path(pdf_path).exists():
        print(f"El archivo PDF ya existe: {pdf_path}")
        continue
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    workbook = excel.Workbooks.Open(filepath)
    try:
        sheet = workbook.Sheets("Hoja 1")
        
        # Configurar la hoja para que se ajuste a una sola página en ancho y alto
        sheet.PageSetup.Zoom = 100

        # Usa orientación horizontal
        sheet.PageSetup.Orientation = 2  # 2 para orientación horizontal

        # Configurar el tamaño del papel A4
        sheet.PageSetup.PaperSize = 8  # A4

        # Ajusta los márgenes
        sheet.PageSetup.LeftMargin = excel.InchesToPoints(0.05) #1.95
        sheet.PageSetup.RightMargin = excel.InchesToPoints(0.05)
        sheet.PageSetup.TopMargin = excel.InchesToPoints(0.05)
        sheet.PageSetup.BottomMargin = excel.InchesToPoints(0.05)

        # Establecer calidad de impresión
        sheet.PageSetup.PrintQuality = 600
        
        # Ajuste final de escala
        sheet.PageSetup.CenterHorizontally = True
        sheet.PageSetup.CenterVertically = True
        
        sheet.ExportAsFixedFormat(0, pdf_path)
    except Exception as e:
        print(f"Error al guardar {filepath} como PDF: {e}")
    finally:
        workbook.Close(SaveChanges=False)
        excel.Quit()


print("Consolidado COMFAMA converted to PDF")