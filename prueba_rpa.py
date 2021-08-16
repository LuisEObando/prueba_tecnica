from selenium import webdriver #Importamos selenium
from selenium.webdriver.common.keys import Keys  #Para realizar acciones por teclado
from selenium.webdriver.common.by import By #nos permite hacer referencia a los elementos para interactuar
from selenium.webdriver.support.ui import WebDriverWait #Para hacer las esperas y validar condiciones
from selenium.webdriver.support import expected_conditions as EC 
import time
import smtplib
from email import encoders ##para codificar los elementos a adjuntar
from email.mime.base import MIMEBase 
from email.mime.text import MIMEText ##mime: (mensaje multimedia)
from email.mime.multipart import MIMEMultipart
from typing import FrozenSet #Ya vienen preinstalados (mensaje multiparte (adjunto, asunto etc))
from openpyxl import load_workbook #Importamos librería que deja cargar excel (solo .xlsx)


def download_zip():
    driver = webdriver.Chrome(executable_path=r'./chromedriver.exe')
    driver.get('http://thedemosite.co.uk/index.php')
    driver.maximize_window()

    try:
        button_zip = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT,'download the free zip here')))
        button_zip.click()
        time.sleep(5)
        
    finally:
        driver.quit() #cerramos la conexión

def send_mail():
    
    #Leemos contraseña de archivo de control no compartido en repositorio
    with open("./control/control_email.txt", "r", encoding="utf-8") as f:
        password_txt = f.readline()
    f.close()
    
    #Accedemos a Excel para leer más datos para el correo
    excel = './datos_destino.xlsx'
    wb = load_workbook(excel)
    sheets = wb.get_sheet_names() #Permite saber cuantas hojas tiene el excel
    sheet1 = wb.get_sheet_by_name('Hoja1')

    #Extraemos los datos
    mail_field_send = sheet1[f'B1'] 
    mail_send = str(mail_field_send.value)
    
    mail_field_dest = sheet1[f'B2']
    mail_dest = str(mail_field_dest.value)

    subject_field = sheet1[f'B3']
    subject_f = str(subject_field.value)

    message_field = sheet1[f'B4']
    message_send = str(message_field.value)

    attachment_field = sheet1[f'B5']
    attachment_send = str(attachment_field.value)

    #iniciamos la conexión con el servidor de correo
    servidor = smtplib.SMTP('smtp.gmail.com', 587) #conectamos al servidor
    servidor.starttls()

    servidor.login(mail_send, str(password_txt)) #accedemos al correo de origen
    
    #Estructura del mensaja a enviar:
    message = MIMEMultipart("alternative") #el correo será un mensaje multiparte de tipo estándar
    message ["Subject"] = subject_f #Pasamos la parte del asunto...
    message ["From"] = mail_send #pasamos el correo de origen
    message ["To"] = mail_dest #pasamos el correo de destino

    #Pasamos algo de html para dar algo de forma al mensaje:
    html = f"""
    <html>
            <head>
                <meta charset="utf-8">
                <title>correo-robot</title>
            </head>
            <body style="background-color: white;">
            <header style = "border: 7px solid rgba(141, 149, 156, 0.637);"">
                <h1 style=" margin: 20px auto; text-align: center; color: rgb(80, 80, 43);">Hola, {str(mail_dest)}</h1>
            </header>
            <section>
                <h2 style="margin: 10px auto; color: rgb(150, 143, 56); text-align: center; ">Prueba RPA</h2>
            </section>
            <section>
                <p style="text-align: justify center ;">{message_send}</p>
                <a style="background-color: rgb(64, 88, 224);
                            display: block;
                            margin-left: auto;
                            margin-right: auto;
                            color: white;
                            padding: 10px;
                            margin: 10px;
                            font-size: 17px;
                            font-family:Cambria, Cochin, Georgia, Times, 'Times New Roman', serif;
                            text-decoration: none;
                            text-transform: uppercase;
                            font-weight: bold;
                            border-radius: 5px;
                            text-align: center;
                            
                            " href="https://github.com/LuisEObando/prueba_tecnica">Visíta Este Repositorio de GitHub Aquí</a>
            </section>
            </body>
            </html>
    """
    part_html = MIMEText(html, "html") #indicamos que es formato html
    message.attach(part_html) #adjuntamos a mensaje la parte html

    #adjuntamos el .zip
    with open(attachment_send, "rb") as attachment: ##lea r como bytes b el archivo
                contenido_adjunto = MIMEBase("application","octet-stream") #Que se interprete como una aplicación (xlsx)
                contenido_adjunto.set_payload(attachment.read())

    encoders.encode_base64(contenido_adjunto) #codificamos el adjunto

    contenido_adjunto.add_header(
    "Content-Disposition",
    f"attachment; filename= {attachment_send}",
    )

    message.attach(contenido_adjunto) #añadimos al mensaje multiparte el adjunto
    correo_empaquetado = message.as_string() #empaquetamos todo el correo

    #Ahora sí, enviamos el correo
    servidor.sendmail(mail_send, mail_dest, correo_empaquetado)
    print('Correo enviado')

    servidor.quit()

def run():
    download_zip()
    send_mail()

if __name__ == '__main__':
    run()