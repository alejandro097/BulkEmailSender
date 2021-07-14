import pandas as pd
import os
import smtplib
import time
from email.message import EmailMessage


#####    Dependencias   ####### 
## pip install pip install xlrd
## pip install openpyxl


## Lista de Correos
e = pd.read_excel("Email.xlsx")
emails = e.values.tolist()

## Datos Email Server
SenderAddress = "TU CORREO"
SenderPassword = "CONTRASENA"


## Ciclo 
count = 1
for i in range(len(emails)):
   receiver = emails[i][0]

   ## Datos Mensaje
   msg = EmailMessage()
   msg['Subject'] = 'Aqui escribir el asunto'
   msg['From'] = SenderAddress
   msg['To'] = receiver
   msg.set_content('Escribir aqui tu mensaje')

   ## Attachment file
   files = ['aqui va nombre del archivo.pdf']
   for file in files:
    with open(file, 'rb') as f:
        file_data = f.read()
        file_name = f.name
   msg.add_attachment(file_data, maintype='application', subtype = 'octet-stream', filename=file_name)

   ## Conectarse al servidor, encriptar la conexion y enviar email
   with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login(SenderAddress, SenderPassword)
    print('\n')
    print('Enviando Correo a ' + receiver)
    smtp.send_message(msg)
    count = count+1
    for i in range(5,0,-1):
        print(f"Esperando {i} segundos para enviar el siguiente correo", end="\r", flush=True)
        time.sleep(1)
        