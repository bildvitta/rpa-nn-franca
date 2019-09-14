from automagica import *
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import datetime
import smtplib
import time

process_arr_data = []

def readInformation(fileName):

	website = 'https://www.franca.sp.gov.br/portal-servico/paginas/publica/processo/consulta.xhtml'

	dataFrame = pd.read_excel(io=fileName)

	rows = dataFrame['ANO'].count()
	year, process, password = '','',''

	#Inside rows of .xls
	for r in range(0,rows):
		#Inside columns
		for s in range(0,6):
			#print (dataFrame.iloc[r,s])
			if (s == 0):
				year = str(dataFrame.iloc[r,s])
				print (year)
			elif (s == 1):
				process = '0' + str(dataFrame.iloc[r,s])
				if (len(process) < 6): 
					process = '0' + process
				print (process)
			elif (s == 2):
				password = str(dataFrame.iloc[r,s])
				print (password)
			else:
				pass

		#Crawl information about selected process.
		crawl(website,year,process,password)
		print ('----------- Done ' + str(r+1) + ' -----------')


#Crawl information about each process.
def crawl(website, year, process, password):

	browser = ChromeBrowser()
	browser.get(website)

	browser.find_element_by_id('formConsulta:anoMask:mskInput').click()
	browser.find_element_by_id('formConsulta:anoMask:mskInput').send_keys(year)

	browser.find_element_by_id('formConsulta:processoMask:mskInput').click()
	browser.find_element_by_id('formConsulta:processoMask:mskInput').send_keys(process)
	
	browser.find_element_by_id('formConsulta:senhaMask:pwdInput').click()
	browser.find_element_by_id('formConsulta:senhaMask:pwdInput').send_keys(password)

	time.sleep(2)
	browser.find_element_by_id('formConsulta:j_idt60').click()

	#print (len(browser.find_elements_by_xpath('//*[@id="formConsulta:j_idt108_data"]/tr')))
	position = (len(browser.find_elements_by_xpath('//*[@id="formConsulta:j_idt108_data"]/tr')))
	#print (position)

	try:
		browser.find_element_by_xpath('//*[@id="formConsulta:j_idt108_data"]/tr[' + str(position) + ']/td/a').click() 
	except:
		print ("Without URL")

	time.sleep(2)

	innerPagePosition = len(browser.find_elements_by_xpath('//*[@id="formConsulta:j_idt143_data"]/tr'))
	#print(innerPagePosition)

	time.sleep(2)

	#Catch the last position, latest move.
	data = browser.find_element_by_xpath('//*[@id="formConsulta:j_idt143_data"]/tr[' + str(innerPagePosition) + ']/td[2]') 
	#print(data.text)

	origin = browser.find_element_by_xpath('//*[@id="formConsulta:j_idt143_data"]/tr[' + str(innerPagePosition) + ']/td[3]') 

	destination = browser.find_element_by_xpath('//*[@id="formConsulta:j_idt143_data"]/tr[' + str(innerPagePosition) + ']/td[4]') 

	manifest = browser.find_element_by_xpath('//*[@id="formConsulta:j_idt143_data"]/tr[' + str(innerPagePosition) + ']/td[5]') 	

	print (data.text,origin.text,destination.text,manifest.text)

	if (checkData(data.text)):
		process_arr_data.append({'Data':data.text, 'Texto':origin.text, 'Destino':destination.text,'Manifestaçaõ':manifest.text})
		print ('Houve alteração e precisa enviar.')

	else:
		print ('Não é igual então não precisa enviar.')

	browser.close()


# Checking if there was a change in processes date.
def checkData (date):
	print (date)
	#print ('Time data: ' + str(time.strftime("%c"))) #Thu Sep  5 19:05:56 2019
	curr_date = datetime.datetime.now()
	curr_format_date = curr_date.strftime("%d/%m/%Y")
	print(curr_format_date) 
	#str(time.strftime("%c")
	return (curr_format_date == date) if True else False


#Connect to SMTP only if is necessary (some process changed).
def connectSMTP ():

	#connect to smpt server
	server = smtplib.SMTP('smtpserver',port)
	server.ehlo()
	server.starttls()
	server.ehlo()

	#from email
	fromaddr = ''
	#to email
	toaddr = ''

	msg = MIMEMultipart()
	msg['From'] = fromaddr
	msg['To'] = toaddr
	msg['Subject'] = 'Relatório de Processos - Incorporação Bild | Vitta Franca'

	#body = 'Ola mundo'
	body = formatText()
	#body = str(process_arr_data) # or plain do a table or something else
	msg.attach(MIMEText(body,'plain'))

	server.login(fromaddr,'yourpasswd')
	server.sendmail(fromaddr,toaddr,msg.as_string())
	server.quit()	

#Format text to send as email body
def formatText ():

	text = ''

	for index_arr in range(len(process_arr_data)):
		for key,val in process_arr_data[index_arr].items():
			text += str(key) + ": " + str(val) + '\n'
		text += '\n\n'
	#print (text)
	return (text)


if __name__ == "__main__":

	#Path to .xls file.
	processTables = 'processos-regional-franca.xlsx'

	#Just to test
	#connectSMTP()
	
	#Initiate looping inside each process from .xls
	readInformation(processTables)

	#If There is any change in processes progress, will be send and e-mail for the responsible.
	if process_arr_data:
		#Connect SMTP and send process update message.
		connectSMTP()
