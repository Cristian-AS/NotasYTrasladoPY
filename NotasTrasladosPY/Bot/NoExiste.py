import pandas as pd
import openpyxl as oxl
import os
import datetime as dt
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from dotenv import load_dotenv
import sys

# Cargar variables de entorno
load_dotenv()

# Obteniendo credenciales de entorno
smtp_server = os.getenv('SMTP_SERVER')
smtp_port = os.getenv('SMTP_PORT')
smtp_username = os.getenv('SMTP_USERNAME')
smtp_password = os.getenv('SMTP_PASSWORD')

smtp_port = smtp_port
email_settings = {'subject': 'Entrega Notas Control de Operaciones $(fecha)', 'recipients': ['aprendiz.funcional@gruporeditos.com', 'Contabilidad@gruporeditos.com', 'monica.zuluaga@gruporeditos.com', 'controldeoperaciones@gruporeditos.com', 'danilo.ortiz@gruporeditos.com']}
msg = """Buenos días,

Lamentamos informar que no se ha encontrado el archivo de Reporte de Notas. Es importante verificar su ubicación en la carpeta correspondiente.

Si tienes alguna pregunta o inquietud sobre la información proporcionada, no dudes en contactarnos a través del siguiente correo, adjuntando la evidencia relevante: correo.automatizacion@gruporeditos.com

Por favor, abstente de responder este correo.

Cordialmente,"""

today = dt.datetime.today().strftime("%d-%m-%Y")

# Configuración del formato del email
message = MIMEMultipart()
message['From'] = smtp_username
message['To'] = ','.join(email_settings.get('recipients'))
subject = email_settings.get('subject').replace('$(fecha)', today)
message["Subject"] = subject
message.attach(MIMEText(msg, 'plain'))

# Enviar el correo electrónico
try:
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(smtp_username, smtp_password)
    server.sendmail(smtp_username, email_settings.get('recipients'), message.as_string())
    server.quit()
    print("Correo electrónico enviado correctamente.")
except Exception as e:
    print(f"Error al enviar el correo electrónico: {e}")
    sys.exit(1)  # Salir del programa con código de error 1