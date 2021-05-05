import time
import configparser
import os
import email
import smtplib
import ssl
import csv

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.message import EmailMessage


WEBDRIVER_PATH = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'chromedriver.exe')
DEFAULT_TIMEOUT = 60
EXPORT_PDF_TIMEOUT = 180
UI_BUG_TIMEOUT = 1


def main():

	web = start_browser()
	web = login(web)
	web, filter_element = load_dashboard(web)
	web = process_report(web, filter_element)

	return web


def load_dashboard(web):
	try:
		print('Esperando la carga de elementos en dashboard.')
		filter_element = WebDriverWait(web, DEFAULT_TIMEOUT).until(expected_conditions.presence_of_element_located((By.CLASS_NAME, 'slicer-dropdown-menu')))
		print('Elementos cargados. Objeto de filtro localizado.')
	except TimeoutException as e:
		print('Error esperando a cargar el dashboard.')
		raise e.__class__()

	return web, filter_element


def process_report(web, filter_element):
	vendors = get_vendors()
	for vendor in vendors:
		print(f"Reporte para {vendor['Representante']}.")
		vendor_filter_name = vendor['Representante'].strip().title()
		filter_element.click()
		try:
			vendor_tag = WebDriverWait(web, DEFAULT_TIMEOUT).until(expected_conditions.element_to_be_clickable((By.XPATH, f"//span[@title='{vendor_filter_name}']")))
		except TimeoutException as e:
			print(f"Error al cargar el elemento de filtro para el vendedor {vendor_filter_name}.")
			raise e.__class__()

		parent_tag = web.find_element_by_xpath(f"//span[@title='{vendor_filter_name}']/ancestor::div[@class='slicerItemContainer']")

		if parent_tag.get_attribute('aria-selected') == 'true':
			pass
		else:
			vendor_tag.click()

		filter_element.click()
		generate_report(web)
		file_path = rename_report(vendor['Representante'])
		send_email(vendor, file_path)
		remove_report(file_path)
		break

	return web


def get_vendors():
	vendors = list()
	vendors_file_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'vendors.csv')
	try:
		with open(vendors_file_path, 'r') as _f:
			reader = csv.DictReader(_f, delimiter=';')
			for v in reader:
				vendors.append(v)
	except Exception as e:
		print('Error al leer archivo de vendedores.')
		raise e.__class__()
	return vendors


def generate_report(web):
	#Export report on current view/filter 
	export_element = web.find_element_by_xpath("//button[@title='Exportar']")
	export_element.click()
	pdf_tag = WebDriverWait(web, DEFAULT_TIMEOUT).until(expected_conditions.presence_of_element_located((By.XPATH, "//span[@localize='Pdf']")))
	pdf_button = WebDriverWait(web, EXPORT_PDF_TIMEOUT).until(expected_conditions.element_to_be_clickable((By.XPATH, "//span[@localize='Pdf']/ancestor::button")))
	pdf_button.click()
	export_button = WebDriverWait(web, DEFAULT_TIMEOUT).until(expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, "button[class='primary'][localize='Export']")))
	export_button.click()
	try:
		close_modal =	web.find_element_by_css_selector("input[class='infonav-dialogCloseIcon']")
		close_modal.click()
	except Exception as e:
		print(e.__class__())
		pass
	print('Espera la notificación.')
	loading_report_box = WebDriverWait(web, EXPORT_PDF_TIMEOUT).until(expected_conditions.presence_of_element_located((By.CSS_SELECTOR, ".spinner")))
	print('Notificación visible.')
	loading_report_box = WebDriverWait(web, EXPORT_PDF_TIMEOUT).until(expected_conditions.invisibility_of_element_located((By.CSS_SELECTOR, ".spinner")))
	print('Se generó el reporte.')
	time.sleep(1)


def rename_report(vendor):
	vendor = get_vendor_file_name(vendor)
	source_file = os.path.join(get_download_folder(), get_config_value('dashboard', 'name')) + '.pdf'
	target_file = os.path.join(get_download_folder(), vendor + '.pdf')
	try:
		os.rename(source_file, target_file)
	except Exception as e:
		remove_report(source_file)
		print('Error al renombrar reporte.')
		raise e.__class__()
	print(target_file)

	return target_file


def remove_report(file_path):
	os.remove(file_path)
	print(f"Removed: {file_path}")


def get_vendor_file_name(text):
	text = text.replace(' ', '_')
	text = text.lower()
	old, new = 'áéíóúñÁÉÍÓÚÑ', 'aeiounAEIOUN'
	trans = str.maketrans(old, new)

	return text.translate(trans)


def start_browser():
	#Config and open browser
	browser_options = Options()
	browser_options.add_argument('start-maximized')
	browser_options.add_argument('--lang=es_PE')
	web = webdriver.Chrome(WEBDRIVER_PATH, chrome_options=browser_options)
	web.get(get_config_value('dashboard', 'url'))

	return web


def login(web):
	mail_element = web.find_element_by_name('loginfmt')
	mail_element.clear()
	mail_element.send_keys(get_config_value('auth', 'email'))
	mail_element.send_keys(Keys.RETURN)
	time.sleep(2)  # cambiar de wait
	passwd_element = web.find_element_by_name('passwd')
	passwd_element.clear()
	passwd_element.send_keys(get_config_value('auth', 'password'))
	passwd_element.send_keys(Keys.RETURN)
	keep_logged_element = web.find_element_by_id('idBtn_Back')
	keep_logged_element.click()

	return web


def get_config_value(section, key):
	parser = configparser.ConfigParser()
	file_config_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'config.cfg')
	parser.read(file_config_path)
	sections = parser.sections()

	if not section in sections:
		raise RuntimeError('La sección de configuración no está definida. Revisar archivo config.cfg')

	return parser[section][key]


def get_download_folder():
	home = os.path.expanduser('~')
	download = os.path.join(home, 'Downloads')

	return download


def send_email(vendor, file_path):
	subject = f"Reporte: {vendor['Representante']}"
	body = ''
	sender_email = get_config_value('email', 'sender')
	receiver_email = 'agocht@gmail.com' #  vendor['Correo']
	try:
		password = get_config_value('email', 'password').replace('"', '')
	except:
		print('No password')
		exit()

	message = EmailMessage()
	message['Subject'] = subject
	message['From'] = sender_email
	message['To'] = ', '.join([receiver_email])
	message.set_content(body)
	context = ssl.create_default_context()

	with open(file_path, 'rb') as _f:
		message.add_attachment(_f.read(), maintype='application', subtype='pdf', filename=vendor['Representante'] + '.pdf')
	
	try:
		with smtplib.SMTP('smtp.office365.com', 587) as server:
			server.ehlo()
			server.starttls(context=context)
			server.login(sender_email, password)
			server.ehlo()
			server.send_message(message)
	except Exception as e:
		remove_report(file_path)
		print('Error al enviar correo. Se elimina el archivo por seguridad.')
		raise e.__class__()

	print(f"Reporte enviado a {vendor['Representante']}.")


if __name__ == '__main__':
	main()