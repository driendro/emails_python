#!/usr/bin/env python3
# -*- coding: iso-8859-1 -*-
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from smtplib import SMTP
import csv

# Obtencion de las lista de destinatarios y archivos
# Deben estar todos dentro de la misma carpeta
# El archivo csv debe estar separado por comas,
# debe contener 2 columnas sin cabecera, 
# archivo.ext y email
a = open('datos.csv')
entrada = csv.reader(a)
registros = list(entrada)
# Cuenta de email y contraseña
user=''
pas=''
# Google
servidor='smtp.gmail.com'
puerto=587
# Hotmail
#servidor='smtp.office365.com'
#puerto=587
asunto='Certificado Seminario Lubricantes'
mensaje= 'Se envia el certificado de asistencia al seminario "Conceptos Basicos de la Lubricacion" realizado en la UTN. Saludos Consejeros de Mecanica'

# Crear una instancia del servidor para envio de correo (hacerlo una sola vez)
# y logeas al servidor
smtp = SMTP(servidor, puerto)
smtp.connect(servidor, puerto)
smtp.ehlo()
smtp.starttls()
smtp.login(user, pas)

for un_reg in registros:
	msg = MIMEMultipart()
	msg['Subject'] = asunto
	msg['From'] = user
	msg['To'] = un_reg[1]
	# Esta es la parte textual:
	part = MIMEText(mensaje)
	msg.attach(part)
	# Esta es la parte binaria (puede ser cualquier extensión):
	part = MIMEApplication(open(un_reg[0],"rb").read())
	part.add_header('Content-Disposition', 'attachment', filename="Certificado.pdf")
	msg.attach(part)
	# Enviar el mail (o los mails)
	smtp.sendmail(msg['From'], msg['To'], msg.as_string())
	print(un_reg)
