from email import encoders
import smtplib
from os.path import join, dirname
import datetime as dt
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from os.path import join
from dotenv import load_dotenv
import os

# Cargar variables de entorno
load_dotenv()

# Obteniendo credenciales de entorno
smtp_server = os.getenv('SMTP_SERVER')
smtp_port = os.getenv('SMTP_PORT')
smtp_username = os.getenv('SMTP_USERNAME')
smtp_password = os.getenv('SMTP_PASSWORD')

# Obtén la ruta del directorio actual donde se encuentra el archivo
current_dir = os.path.dirname(os.path.abspath(__file__))

# Mueve un nivel hacia arriba
workfolder_path = os.path.dirname(current_dir)

# Variables de configuración
smtp_port = smtp_port
email_settings = {'subject': 'Entrega Notas Control de Operaciones $(fecha)', 'recipients': ['aprendiz.funcional@gruporeditos.com', 'Contabilidad@gruporeditos.com', 'monica.zuluaga@gruporeditos.com', 'controldeoperaciones@gruporeditos.com', 'danilo.ortiz@gruporeditos.com']}
msg = """Buenos días,

Mediante el proceso automático RPA se hace entrega a Contabilidad de las notas creadas
por canales como se relaciona en el archivo adjunto, así mismo, se envían los soportes en
PDF de las notas en mención.
 
Cualquier duda o inquietud con la información reportada contacta al siguiente correo
adjuntando la evidencia correspondiente correo.automatizacion@gruporeditos.com
 
NOTA: Por favor seguir el siguiente procedimiento para la extracción de la información,
paso 1: Descargar el .zip,  paso 2: dar click derecho al archivo, paso 3: seleccionar la opción 7zip y  paso 4: seleccionar extraer aquí(extract here).
 
Por favor no responder ni enviar correos de respuesta a la cuenta 
automata.mercadeo@gruporeditos.com.
 
Cordialmente,"""

# Declaración de una variable para la fecha de hoy
today = dt.datetime.today().strftime("%d-%m-%Y")

# Configuración del formato del email
message = MIMEMultipart()
message['From'] = smtp_username
message['To'] = ','.join(email_settings.get('recipients'))
subject = email_settings.get('subject').replace('$(fecha)', today)
message["Subject"] = subject
message.attach(MIMEText(msg, 'plain'))

# Rutas de los archivos adjuntos
folder = join(workfolder_path, "reports", today)
zip_path = join(folder, 'Resultado.zip')
report_path = join(folder, f'CONSOLIDADO {today}.xlsx')

# Adjuntar el archivo ZIP al mensaje
with open(zip_path, mode='rb') as part:
    zip_file = MIMEBase("application", "zip")
    zip_file.set_payload(part.read())
encoders.encode_base64(zip_file)
zip_file.add_header("Content-Disposition", "attachment", filename="Resultado.zip")
message.attach(zip_file)

# Adjuntar el archivo Excel consolidado al mensaje
with open(report_path, mode='rb') as part:
    excel_file = MIMEBase('application', 'octet-stream')
    excel_file.set_payload(part.read())
encoders.encode_base64(excel_file)
excel_file.add_header('Content-Disposition', 'attachment', filename='Consolidado.xlsx')
message.attach(excel_file)

# Envío de correo mediante una conexión SMTP
try:
    with smtplib.SMTP(host=smtp_server, port=smtp_port, timeout=60) as conn:
        conn.starttls()
        conn.login(user=smtp_username, password=smtp_password)
        conn.sendmail(from_addr=smtp_username, to_addrs=email_settings.get('recipients'), msg=message.as_string())
except Exception as e:
    print(f"The connection has thrown an error: {e}")
else:
    print("The message has been sent successfully")
