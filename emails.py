from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from smtplib import SMTP
from jinja2 import Template
import pandas as pd

# Alias
# Alias de correo, normalemnte es el mismo que el smtp_user
alias = 'Juan Perez <juan.perez@dominio.com>'
# el correo que usas para enviar los mensajes
smtp_user = 'juan.perez@dominio.com'
smtp_pass = 'contraseña_correo'  # password de la cuenta de correo

# Taplate en html, usando jinja2
asunto = 'De q trata el correo'
archivo_template = 'html/template.html'
db_csv = 'archivo.csv'

def envio_correo(i, j, nombre, apellido, correo1, correo2):
    msg['To'] = '%s' % (correo1)
    msg['Cc'] = '%s' % (correo2)
    try:
        smtp.sendmail(msg['From'], (msg['To'], msg['Cc']), msg.as_string())
        s = (i, j, nombre, apellido, 'ok', correo1, correo2)
    except:
        smtp.close()
        print(i, correo1)
        exit()
    return(s)


# Obtencion de las lista de destinatarios y archivos
# Deben estar todos dentro de la misma carpeta

registros = pd.read_csv(db_csv, delimiter=';')

# Crear una instancia del servidor para envio de correo (hacerlo una sola vez)
smtp = SMTP("smtp.office365.com", 587)

i = 0  # contador de registros
j = len(registros)  # cantidad de registros
k = 0  # contador de envios por ciclo
k_max = 10  # maximo de envios por ciclo

while i < j:
    if k == 0:
        print('abrir coneccion')
        # Iniciar sesión en el servidor (si es necesario):
        smtp.connect("smtp.office365.com", 587)
        smtp.ehlo()
        smtp.starttls()
        # Tu usuario y tu contraseña
        smtp.login(smtp_user, smtp_pass)

    k = k+1
    #apellido, nombre = str(registros.loc[i, 'nombre']).split(',')
    nombre = str(registros.loc[i, 'Nombre']).title().strip()
    #apellido = str(registros.loc[i, 'Apellido']).upper().strip()
    apellido = ''
    ### Si tiene mas de un correo, tomamos los 2 y enviamos a ambos el mensaje
    ### En el caso de q uno este vacio no se envia nada
    correo1 = str(registros.loc[i, 'correo1'])
    if correo1 != '' and correo1 != 'nan':
        correo1 = correo1.lower().strip()
    correo2 = str(registros.loc[i, 'correo2'])
    #correo2 = str('')
    if correo2 != '' and correo2 != 'nan':
        correo2 = correo2.lower().strip()
    msg = MIMEMultipart()
    msg['Subject'] = asunto
    msg['From'] = alias
    # Armado del mensjae con jinga desde el template
    temp = open(archivo_template, encoding="utf-8").read()
    dic = {
        'nombre': nombre,
        'apellido': apellido
    }
    mensage = Template(temp).render(dic)

    # Esta es la parte textual:
    part = MIMEText(mensage, 'html')
    msg.attach(part)

    # Esta es la imagen a embeber
    #fp = open('template/photo.jpeg', 'rb')
    #flayer = MIMEImage(fp.read())
    #fp.close()
    #fp = open('template/QR.png', 'rb')
    #qr = MIMEImage(fp.read())
    #fp.close()
    ## Define the image's ID as referenced above
    #flayer.add_header('Content-ID', '<flayer>')
    #qr.add_header('Content-ID', '<qr>')
    #msg.attach(qr)
    #msg.attach(flayer)

    # Esta es la parte binaria (puede ser cualquier extensión):
    #archivo='becaypf.png'
    #part = MIMEApplication(open("%s" %(archivo),"rb").read())
    #part.add_header('Content-Disposition', 'attachment', filename="BecaYPF.png")
    #msg.attach(part)
    # Enviar el mail (o los mails) a grupos especificos

    if True:  # int(registros.loc[i, 'id_presinc']):# > 3974:
        s = envio_correo(i, j, nombre, apellido, correo1, correo2)
        print(s)

    # Cada 10 correos, reinicia la coneccion al SMTP
    if k == k_max:
        print('cerrar coneccion')
        smtp.close()
        k = 0
    elif i == j-1:
        print('cerrar coneccion')
        smtp.close()
        k = 0
    i = i+1
