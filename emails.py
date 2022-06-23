from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from smtplib import SMTP
from jinja2 import Template
import pandas as pd

# Alias
# Alias de correo, normalemnte es el mismo que el smtp_user
alias = 'correo@frlp.utn.edu.ar'
# el correo que usas para enviar los mensajes
smtp_user = 'correo@frlp.utn.edu.ar'
smtp_pass = 'Contraseña'  # password de la cuenta de correo


# Taplate en html, usando jinja2
asunto = 'Asunto del email'
archivo_template = 'path/al/template.html'


def envio_correo(i, j, nombre, apellido, correo):
    msg['To'] = '%s' % (correo)
    try:
        smtp.sendmail(msg['From'], msg['To'], msg.as_string())
        s = (i, j, nombre, apellido, 'ok', correo)
    except:
        smtp.close()
        k = 0
        s = 'Error, reintentar'
        print(i, correo)
        exit()
    return(s)


# Obtencion de las lista de destinatarios y archivos
# Deben estar todos dentro de la misma carpeta
registros = pd.read_csv('path/al/archivo.csv', delimiter=';')

# para testear, sobreescribe la lista registros por un dado
# comentar para enviar el correo a los desinatarios
# armar el diccionario con las columnas de CSV
#dic= {
#    'Titulo':'',
#    'Nombre':'',
#    'Apellido':'',
#    'correo1': '',
#    'correo2': ''
#    }
#registros=pd.DataFrame(dic,index=[0])

# Crear una instancia del servidor para envio de correo (hacerlo una sola vez)
smtp = SMTP("smtp.office365.com", 587)

i = 180  # contador de registros
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
    apellido = str(registros.loc[i, 'Apellido']).title().strip()
    ### Si tiene mas de un correo, tomamos los 2 y enviamos a ambos el mensaje
    ### En el caso de q uno este vacio no se envia nada
    correo1 = str(registros.loc[i, 'correo1'])
    if correo1 != '':
        correo1 = correo1.lower().strip()
    correo2 = str(registros.loc[i, 'correo2'])
    if correo2 != '':
        correo2 = correo2.lower().strip()
    msg = MIMEMultipart()
    msg['Subject'] = asunto
    msg['From'] = alias
    # Armado del mensjae con jinga desde el template
    temp = open(archivo_template).read()
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
    s1 = '()'
    s2 = '()'
    if correo1 != '' and correo1 != 'nan':
        s1 = envio_correo(i, j, nombre, apellido, correo1)
    if correo2 != '' and correo2 != 'nan':
        s2 = envio_correo(i, j, nombre, apellido, correo2)
    print(s1, s2)

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
