from dotenv import load_dotenv
from automagica import ChromeBrowser
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from oauth2client.service_account import ServiceAccountCredentials

import os
import time
import smtplib
import gspread
from datetime import datetime

# Loading .env file variables.
load_dotenv()

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name(
    os.getenv('GS_CREDENTIALS_PATH'),
    scope
)

gc = gspread.authorize(credentials)

# Open a worksheet from spreadsheet with one shot
wks = gc.open_by_key(
    os.getenv('GS_SPREADSHEET_ID')
).worksheet(
    os.getenv('GS_WKS_NAME')
)

# Fetch all processes
process_list = wks.get_all_values()

# Updated processes.
process_arr_data = []

# Selenium instance
browser = ChromeBrowser()

# E-mail address list to send notifications to.
email_to_address = ['estevao.simoes@bild.com.br']


def readInformation():

    website = 'https://www.franca.sp.gov.br/portal-servico/paginas/publica/processo/consulta.xhtml'

    index = 1

    for process in process_list[1:]:

        if str(process[6]) == 'SIM':
            # Crawl information about selected process.
            crawl(website, process[0], process[1],
                  process[2], process[4], index)
            print('----------- Done -----------')
        else:
            print('----------- Ignoring process -----------')
        index = index + 1
    return browser.quit()


'''
Crawl information about each process.
'''


def crawl(website, year, process, password, project_name, line_number):

    browser.get(website)

    browser.find_element_by_id('formConsulta:anoMask:mskInput').click()
    browser.find_element_by_id('formConsulta:anoMask:mskInput').send_keys(year)

    browser.find_element_by_id('formConsulta:processoMask:mskInput').click()
    browser.find_element_by_id(
        'formConsulta:processoMask:mskInput').send_keys(process)

    browser.find_element_by_id('formConsulta:senhaMask:pwdInput').click()
    browser.find_element_by_id(
        'formConsulta:senhaMask:pwdInput').send_keys(password)

    # submitting form
    browser.find_element_by_id('formConsulta:j_idt60').click()

    # validando se o processo existe.
    try:

        browser.find_element_by_xpath(
            '//*[@id="formConsulta:messages"]/div/ul/li/span').text
        print('processo não existe.')

        wks.update_acell('I' + str(line_number + 1),
                         datetime.now().strftime('%d-%m-%Y %H:%M:%S'))
        wks.update_acell('J' + str(line_number + 1), 'ERRO')

        return
    except:
        wks.update_acell('I' + str(line_number + 1),
                         datetime.now().strftime('%d-%m-%Y %H:%M:%S'))
        wks.update_acell('J' + str(line_number + 1), 'OK')
        print('processo existe, continuando.')

    position = (len(browser.find_elements_by_xpath(
        '//*[@id="formConsulta:j_idt108_data"]/tr')))

    try:
        process_name = (str(browser.find_element_by_xpath(
            '//*[@id="formConsulta:cabecalhoProcesso_content"]/table/tbody/tr[3]/td/label').text))
        print(process_name)
    except:
        print('Error in find process name')

    try:
        browser.find_element_by_xpath(
            '//*[@id="formConsulta:j_idt108_data"]/tr[' + str(position) + ']/td/a').click()
    except:
        print("Single page process. Going to the next one.")

    # Slowing down so it does not crash.
    time.sleep(0.8)

    innerPagePosition = len(browser.find_elements_by_xpath(
        '//*[@id="formConsulta:j_idt143_data"]/tr'))

    # Catch the last position, latest move.
    data = browser.find_element_by_xpath(
        '//*[@id="formConsulta:j_idt143_data"]/tr[' + str(innerPagePosition) + ']/td[2]')

    origin = browser.find_element_by_xpath(
        '//*[@id="formConsulta:j_idt143_data"]/tr[' + str(innerPagePosition) + ']/td[3]')

    destination = browser.find_element_by_xpath(
        '//*[@id="formConsulta:j_idt143_data"]/tr[' + str(innerPagePosition) + ']/td[4]')

    manifest = browser.find_element_by_xpath(
        '//*[@id="formConsulta:j_idt143_data"]/tr[' + str(innerPagePosition) + ']/td[5]')

    print(data.text, origin.text, destination.text, manifest.text)

    if (checkDate(data.text)):
        process_arr_data.append({'Ano': year, 'Nome do projeto': project_name, 'Protocolo': process, 'Nome do Processo': process_name,
                                 'Processo': password, 'Data': data.text, 'Origem': origin.text, 'Destino': destination.text, 'Manifestação': manifest.text})
        print(str(process) + ': houve alteração')
        wks.update_acell('H' + str(line_number + 1),
                         datetime.now().strftime('%d-%m-%Y %H:%M:%S'))

    else:
        print(str(process) + ': sem alteração')


# Compare process date to current date
def checkDate(date):

    curr_format_date = datetime.now().strftime("%d/%m/%Y")

    return (curr_format_date == date) if True else False


# Connect to SMTP and send noticiatino e-mail
def sendNotificationEmail():

    server = smtplib.SMTP(os.getenv('SMTP_HOST'), os.getenv('SMTP_PORT'))

    server.ehlo()
    server.starttls()
    server.ehlo()

    server.login(os.getenv('SMTP_USER'), os.getenv('SMTP_PASS'))

    msg = MIMEMultipart()
    msg['From'] = os.getenv('SMTP_MAIL_FROM')
    msg['To'] = ', '.join(email_to_address)
    msg['Subject'] = 'Relatório de Acompanhamento RPA - Incorporação Bild & Vitta Franca'

    # Check if theres any process updates to send
    if process_arr_data:
        body = formatText()
    else:
        body = 'Eu sou o JARVIS, escaneei o site da Prefeitura para o senhor(a) e não encontrei nada de novo!'

    msg.attach(MIMEText(body, 'plain'))

    server.sendmail(os.getenv('SMTP_MAIL_FROM'),
                    email_to_address, msg.as_string())

    return server.quit()


def formatText():

    text = 'Eu sou o JARVIS, escaneei o site da Prefeitura para o senhor(a) e olha os trâmites que encontrei:\n\n'

    for index_arr in range(len(process_arr_data)):
        for key, val in process_arr_data[index_arr].items():
            text += str(key) + ": " + str(val) + '\n'
        text += '\n\n'
    print(text)
    return (text)


if __name__ == "__main__":

    readInformation()

    print('Tentando enviar e-mail.')

    # Trying to send e-mail
    try:
        sendNotificationEmail()
        print('Email enviado.')
    except smtplib.SMTPConnectError:
        print('Não foi possível conectar no servidor de e-mail')
    except smtplib.SMTPAuthenticationError:
        print('Usuário e senha de e-mail incorretos.')
    except smtplib.SMTPRecipientsRefused:
        print('Destinatário inválido.')
