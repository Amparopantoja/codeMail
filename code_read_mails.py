import imaplib
import email
from email.header import decode_header
import msvcrt
import os
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import pandas as pd
import re
import json 

with open('config.json') as config_file:
    config_data = json.load(config_file)
    
    
ARBA_MAIL = config_data['ARBA_MAIL']
AGIP_MAIL = config_data['AGIP_MAIL']
TUCUMAN_MAIL = config_data['TUCUMAN_MAIL']
CORDOBA_MAIL = config_data['CORDOBA_MAIL']
AFIP_MAIL = config_data['AFIP_MAIL']
USERNAME_MAIL = config_data['USERNAME_MAIL']
PWD_MAIL = config_data['PWD_MAIL']
FILE_PATH = "C:/Users/Amparo/Desktop/CODEMAILS/Asesoramiento Contable/Planillas WNS Unificadas/Seguimiento temas comunes/DFE/Notificaciones/2023/Septiembre/"
FILENAME = FILE_PATH + "Notificaciones DFE.xlsx"
SHEETNAME_BD = "BD"
SHEETNAME_ABM = "ABM"
SHEETNAME_RESPONSABLE = "RESPONSABLES"
SHEETNAME_LIDER = "LIDERES"

    
def base_writer(hoja, index, col, value):
    celda = hoja[col + str(index)]
    celda.value = value
    
def get_organismo(b, f):
    if "ARBA" in b.upper():
        return "ARBA"
    elif AFIP_MAIL == f:
        return "AFIP"
    elif AGIP_MAIL == f:
        return "AGIP"
    elif TUCUMAN_MAIL == f:
        return "TUCUMAN"
    elif CORDOBA_MAIL == f:
        return "CORDOBA"
    else:
        return ""
    
def get_cuit(b, f):
    if "ARBA" in b.upper():
        return b.split('en el contribuyente ')[1]
    elif AFIP_MAIL == f:
        return b.split('para el cuit ')[1].split('.')[0].replace('-','')
    elif AGIP_MAIL == f:
        return b.split('CUIT: ')[1]
    elif TUCUMAN_MAIL == f:
        return b.split('contribuyente ')[1].split(' ')[0]
    elif CORDOBA_MAIL == f:
        return b.split('contribuyente ')[1].split(' ')[0]
    else:
        return ""

def get_tema(b, f, s):
    if "ARBA" in b.upper():
        return ""
    elif AFIP_MAIL == f:
        return s.split('"')[1]
    elif AGIP_MAIL == f:
        return ""
    elif TUCUMAN_MAIL == f:
        return ""
    elif CORDOBA_MAIL == f:
        return ""
    else:
        return ""
      
def show_menu():
    while(True):
        os.system('cls')
        print("1 - cargar mails")
        print("2 - descargar archivos")
        print("3 - enviar mails")
        print("4 - salir")
        opcion = int(input())
        
        if opcion == 1:
            os.system('cls')
            load_mails()
            print("mails cargados")
            # Imprime un mensaje
            print("Presiona cualquier tecla para continuar...")

            # Espera a que se presione una tecla
            msvcrt.getch()
        elif opcion == 2:
            print("descargar archivos")  
            print("Presiona cualquier tecla para continuar...") 
            msvcrt.getch()
        elif opcion == 3:
            print("enviar mails")
            print("Presiona cualquier tecla para continuar...")
            msvcrt.getch()
        elif opcion == 4:
            break
        
def load_mails():
    lista_de_correos = []
    
    df = pd.read_excel(FILENAME, sheet_name=SHEETNAME_ABM)
    df_responsable = pd.read_excel(FILENAME, sheet_name=SHEETNAME_RESPONSABLE)
    df_lider = pd.read_excel(FILENAME, sheet_name=SHEETNAME_LIDER)


    # Conéctate al servidor IMAP de Gmail
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(USERNAME_MAIL, PWD_MAIL)

    # Selecciona la bandeja de entrada (inbox)
    imap.select("inbox")
    # Busca los últimos 10 correos electrónicos
    status, message_ids = imap.search(None, "ALL")
    message_ids = message_ids[0].split()

    # Itera sobre los últimos 10 correos electrónicos
    for message_id in message_ids[-2:]:
        status, msg_data = imap.fetch(message_id, "(RFC822)")
        raw_email = msg_data[0][1]
        email_message = email.message_from_bytes(raw_email)
        subject, encoding = decode_header(email_message["Subject"])[0]
        codificacion = email_message.get_content_charset()
        
            # Extrae el contenido del correo electrónico
        if email_message.is_multipart():
            for part in email_message.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

                if "attachment" not in content_disposition:
                    payload = part.get_payload(decode=True)
                    if payload is not None:
                        body = payload.decode()  # Intenta decodificar el contenido
                    else:
                        body = part.get_payload()  # Si no se puede decodificar, obtén el contenido sin decodificar
                        #urls = re.findall(r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', body)

        else:
            body = email_message.get_payload(decode=True).decode(codificacion, "ignore")
            #urls = re.findall(r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', body)


        imap.store(message_id, "-FLAGS", "(\Seen)")

        imap.store(message_id, '+X-GM-LABELS', "label")

        if isinstance(subject, bytes):
            subject = subject.decode(encoding or "utf-8")

        lista_de_correos.append(Correo(subject, email_message['From'],email_message['Date'], body))


    # Abre el archivo Excel
    workbook = openpyxl.load_workbook(FILENAME)
    sheet = workbook[SHEETNAME_BD]
   
    for c in lista_de_correos:
        #Categorizar email
        #Escribir en el excel
        sheet.insert_rows(2)
    
        base_writer(sheet, 2, 'A',c.date)
        organismo = get_organismo(c.body, c.fromMail)
        base_writer(sheet, 2, 'B', organismo)
        cuit = get_cuit(c.body, c.fromMail)
        #if cuit != '':
            
        base_writer(sheet, 2, 'C', cuit)
        
        tema = get_tema(c.body, c.fromMail, c.subject)
        base_writer(sheet, 2, 'K', tema)
        
        for i, f in df.iterrows():
            if str(f.CUIT) == cuit:
                cliente = f.Contribuyente
                responsable = f.Responsable
                equipo = f.Equipo
                for r in df_responsable.iterrows():
                    if r[1].Nombre == responsable:
                        mail_responsable = r[1].Mail
                
                for r in df_lider.iterrows():
                    if r[1].Nombre == equipo:
                        mail_copia = r[1].Mail
                cliente_sueldo = f["CLIENTE SUELDOS"]
                
                base_writer(sheet, 2, 'D', cliente)
                base_writer(sheet, 2, 'E', responsable)
                base_writer(sheet, 2, 'F', equipo)
                base_writer(sheet, 2, 'G', mail_responsable)
                base_writer(sheet, 2, 'H', mail_copia)
                base_writer(sheet, 2, 'I', cliente_sueldo)

        
        print(f"Subject: {c.subject}")
        print(f"From: {c.fromMail}")
        print(f"Date: {c.date}")
        print(f"Body: {c.body}")
        print("=" * 50)
    
    workbook.save(FILENAME)
    workbook.close()
    # Cierra la conexión
    imap.logout()

        
class Correo:
    def __init__(self, subject, fromMail, date, body):
        # Atributos de la clase
        self.subject = subject
        self.fromMail = fromMail
        self.date = date
        self.body = body


show_menu()