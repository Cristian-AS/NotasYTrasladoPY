from os.path import join
from os import listdir
import datetime as dt
import win32com.client
from pathlib import Path
import os

# Obtén la ruta del directorio actual donde se encuentra el archivo
current_dir = os.path.dirname(os.path.abspath(__file__))

# Mueve un nivel hacia arriba
workfolder_path = os.path.dirname(current_dir)

# Nombre de la carpeta donde serán almacenados los reportes generados.
folder_name = 'notas'

# Declaración de una variable para la fecha de hoy
today = dt.datetime.today().strftime("%d-%m-%Y")

# ruta de la carpeta donde estan almacenados los reportes a convertir.
folder = join(workfolder_path, "reports", today, folder_name)

files = [file.split('.')[0] for file in listdir(folder)]

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
        sheet = workbook.Sheets("Hoja1")
        
        # Configurar la hoja para que se ajuste a una sola página en ancho y alto
        sheet.PageSetup.Zoom = 100

        # Usa orientación horizontal
        sheet.PageSetup.Orientation =  2 # 2 para orientación horizontal

        # Configurar el tamaño del papel A4
        sheet.PageSetup.PaperSize = 7  # A4

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
    
print("All notas converted to pdf")
