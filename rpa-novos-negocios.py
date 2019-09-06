from automagica import *
import pandas as pd 
import time

'''
Read information from .xls and interactively use crawl() function to get information about each one.

Pandas read excel: https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.read_excel.html

'''
def readInformation(fileName):

	website = 'https://www.franca.sp.gov.br/portal-servico/paginas/publica/processo/consulta.xhtml'


	dataFrame = pd.read_excel(io=fileName)
	#print (dataFrame.head(5))
	#print (dataFrame.head())
	#print (dataFrame.iloc[0,0])
	#print(dataFrame['ANO'].count()) #Number of rows in excel.
	rows = dataFrame['ANO'].count()
	year, process, password = '','',''

	#Inside rows
	for r in range(6,rows):
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
		print ('----------- Done ' + str(r) + ' -----------')

'''
Crawl information about each process.
'''
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

	#print(type(data.text))
	if (checkData(data.text)):
		# SE A DATA FOR IGUAL A DO DIA, ENVIAR NOTIFICAÇÃO
		print ('Via e-mail')
	else:
		print ('Equal')

	browser.close()

def checkData (data):
	print (data)
	print ('Time data: ' + str(time.strftime("%c"))) #Thu Sep  5 19:05:56 2019
	return (str(time.strftime("%c")) == data) if print("It's Equal") else print("It's Equal")


if __name__ == "__main__":

	processTables = 'processos-regional-franca.xlsx'
	
	readInformation(processTables)