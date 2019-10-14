# -*- coding: iso-8859-1 -*-
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from smtplib import SMTP
import csv

# Obtencion de las lista de destinatarios y archivos
# Deben estar todos dentro de la misma carpeta
a = open('datos.csv')
entrada = csv.reader(a)
registros = list(entrada)

# Crear una instancia del servidor para envio de correo (hacerlo una sola vez)
smtp = SMTP("smtp.gmail.com", 587)
# Iniciar sesión en el servidor (si es necesario):
smtp.connect("smtp.gmail.com", 587)
smtp.ehlo()
smtp.starttls()
smtp.login("tu@email.com", "tu_pass") # Tu usuario y tu contraseña

for un_reg in registros:
	msg = MIMEMultipart()
	msg['Subject'] = 'Certificado de Gil'
	msg['From'] = 'quien_envia@email.com'
	msg['To'] = '%s' %(un_reg[1])
	# Esta es la parte textual:
	part = MIMEText("Hola %s, %s: te paso un archivo interesante" %(un_reg[2], un_reg[3]))
	msg.attach(part)
	# Esta es la parte binaria (puede ser cualquier extensión):
	part = MIMEApplication(open("%s" %(un_reg[0]),"rb").read())
	part.add_header('Content-Disposition', 'attachment', filename="Certificado.pdf")
	msg.attach(part)
	# Enviar el mail (o los mails)
	smtp.sendmail(msg['From'], msg['To'], msg.as_string())
	print(un_reg)
 