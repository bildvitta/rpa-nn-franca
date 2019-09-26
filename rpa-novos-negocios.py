from automagica import ChromeBrowser
from dotenv import load_dotenv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import os
import time
import smtplib
import datetime
import pandas as pd

# Loading .env file variables.
load_dotenv()

# Updated processes.
process_arr_data = []

# Selenium instance
browser = ChromeBrowser()

# E-mail address list to send notifications to.
email_to_address = ['']


def readInformation(fileName):

    website = 'https://www.franca.sp.gov.br/portal-servico/paginas/publica/processo/consulta.xhtml'

    dataFrame = pd.read_excel(io=fileName)
    rows = dataFrame['ANO'].count()
    year, process, password, process_name, project_name = '', '', '', '', ''

    # Inside rows of .xls
    for r in range(0, rows):
        # Inside columns
        for s in range(0, 6):
            if (s == 0):
                year = str(dataFrame.iloc[r, s])
                print(year)
            elif (s == 1):
                process = '0' + str(dataFrame.iloc[r, s])
                if (len(process) < 6):
                    process = '0' + process
                print(process)
            elif (s == 2):
                password = str(dataFrame.iloc[r, s])
                print(password)
            elif (s == 4):
                project_name = str(dataFrame.iloc[r, s])
                print(project_name)
            else:
                pass

        # Crawl information about selected process.
        crawl(website, year, process, password, project_name)
        print('----------- Done ' + str(r+1) + ' -----------')

    return browser.quit()


'''
Crawl information about each process.
'''


def crawl(website, year, process, password, project_name):

    browser.get(website)

    browser.find_element_by_id('formConsulta:anoMask:mskInput').click()
    browser.find_element_by_id('formConsulta:anoMask:mskInput').send_keys(year)

    browser.find_element_by_id('formConsulta:processoMask:mskInput').click()
    browser.find_element_by_id(
        'formConsulta:processoMask:mskInput').send_keys(process)

    browser.find_element_by_id('formConsulta:senhaMask:pwdInput').click()
    browser.find_element_by_id(
        'formConsulta:senhaMask:pwdInput').send_keys(password)

    browser.find_element_by_id('formConsulta:j_idt60').click()

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
                                 'Processo': password, 'Data': data.text, 'Texto': origin.text, 'Destino': destination.text, 'Manifestação': manifest.text})
        print(str(process) + ': houve alteração')

    else:
        print(str(process) + ': sem alteração')


# Compare process date to current date
def checkDate(date):

    curr_date = datetime.datetime.now()
    curr_format_date = curr_date.strftime("%d/%m/%Y")

    return (curr_format_date == date) if True else False


# Connect to SMTP and send noticiatino e-mail
def sendNotificationEmail():

    server = smtplib.SMTP(os.getenv('SMTP_HOST'), os.getenv('SMTP_PORT'))

    server.ehlo()
    server.starttls()
    server.ehlo()

    server.login(os.getenv('SMTP_USER'), os.getenv('SMTP_PASS'))

    msg = MIMEMultipart()
    msg['From'] = os.getenv('MAIL_FROM')
    msg['To'] = ', '.join(email_to_address)
    msg['Subject'] = 'Relatório de Acompanhamento RPA - Incorporação Bild & Vitta Franca'

    # Check if theres any process updates to send
    if process_arr_data:
        body = formatText()
    else:
        body = 'Olá, tudo bem? Eu sou o robozinho que veio ajudar na verificação do andamento dos processos.\nNão houveram alterações em projetos até o momento!'

    msg.attach(MIMEText(body, 'plain'))

    server.sendmail(os.getenv('MAIL_FROM'), email_to_address, msg.as_string())

    return server.quit()


def formatText():

    text = 'Olá, tudo bem? Eu sou o robozinho que veio ajudar na verificação do andamento dos processos. Segue abaixo os protocolos que foram atualizados recentemente!\n\n\n'

    for index_arr in range(len(process_arr_data)):
        for key, val in process_arr_data[index_arr].items():
            text += str(key) + ": " + str(val) + '\n'
        text += '\n\n'
    print(text)
    return (text)


if __name__ == "__main__":

    # Complete path to .xls file, so it works thru a bat file.
    processTables = 'C:\\Users\\DELL\\Documents\\projetos\\rpa-nn\\processos-regional-franca.xlsx'

    readInformation(processTables)

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
