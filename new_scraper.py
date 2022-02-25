# 02/2020 Agustín Escobar

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from tribunales import TRIBUNAL_CODES
from selenium.webdriver.common.by import By
from openpyxl import Workbook
from datetime import date
import json
import smtplib
from os.path import basename, exists
from os import remove
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate

successList = [
  'Proc.: Designación Árbitro',
  'Proc.: Voluntario - Notificaciones judiciales'
]

workBook = Workbook()
page = workBook.active
page.append(('ROL', 'TRIBUNAL', 'INGRESO', 'ESTADO', 'URL'))

options = Options()
options.add_argument("--headless")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("enable-automation")
options.add_argument("--disable-infobars")
options.add_argument("--disable-dev-shm-usage")
web = webdriver.Chrome(options=options)

retryDict = {}

for tribunal_code in TRIBUNAL_CODES:
  web.get('https://civil.pjud.cl/CIVILPORWEB/')
  try:
    f = web.find_element(By.NAME, 'body')
  except:
    retryDict[tribunal_code] = f'Tribunal entero ({TRIBUNAL_CODES[tribunal_code]})'
    continue
  web.switch_to.frame(f)

  dateTabWeb = web.find_element(By.XPATH, '//*[@id="tdDos"]')
  dateTabWeb.click()

  web.find_element(By.XPATH, '//*[@id="IMG_FEC_Desde"]').click()

  monthWeb = Select(web.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr/td/center/table[1]/tbody/tr/td[1]/select')).select_by_value(f'{date.today().month}')
  yearWeb = Select(web.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr/td/center/table[1]/tbody/tr/td[3]/select')).select_by_value(f'{date.today().year}')
  startDaysWeb = web.find_elements(By.CLASS_NAME, 'TESTcpCurrentMonthDate')
  for day in startDaysWeb:
    if day.text == f'{date.today().day - 1}' and day.tag_name == 'a':
      day.click()
      break

  web.find_element(By.XPATH, '//*[@id="IMG_FEC_Hasta"]').click()

  monthWeb = Select(web.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr/td/center/table[1]/tbody/tr/td[1]/select')).select_by_value(f'{date.today().month}')
  yearWeb = Select(web.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr/td/center/table[1]/tbody/tr/td[3]/select')).select_by_value(f'{date.today().year}')
  startDaysWeb = web.find_elements(By.CLASS_NAME, 'TESTcpCurrentMonthDate')
  for day in startDaysWeb:
    if day.text == f'{date.today().day - 1}' and  day.tag_name == 'a':
      day.click()
      break

  try:
    tribunalWeb = Select(web.find_element(By.XPATH, '//*[@id="tribUno"]/select')).select_by_value(tribunal_code)
  except:
    retryDict[tribunal_code] = f'Error en {tribunal_code}, {yearWeb}, {monthWeb}'
    continue
  submitWeb = web.find_element(By.XPATH, '/html/body/form/table[6]/tbody/tr/td[2]/a[1]').click()
  try:
    bigTableWeb = web.find_element(By.XPATH, '//*[@id="contentCellsAddTabla"]')
  except:
    retryDict[tribunal_code] = f'Error en {submitWeb}'
    continue
  searchArray = []
  for element in bigTableWeb.find_elements(By.CLASS_NAME, 'textoC'):
    if '/' not in element.text:
      if 'C' in element.text: 
        searchArray.append(element.find_element(By.TAG_NAME, 'a').get_attribute('href'))
      if 'V' in element.text:
        searchArray.append(element.find_element(By.TAG_NAME, 'a').get_attribute('href'))

  web.execute_script("window.open('','_blank');")
  web.switch_to.window(web.window_handles[1])
  for link in searchArray:
    try:
      web.get(link)
      causeWeb = web.find_element(By.XPATH, '/html/body/form/table[3]/tbody/tr[2]/td[2]')
      if causeWeb.text in successList:
        rol = web.find_element(By.XPATH, '/html/body/form/table[3]/tbody/tr[1]/td[1]').text.split(' ')[2]
        tribunal = web.find_element(By.XPATH, '/html/body/form/table[3]/tbody/tr[4]/td[1]').text.split(' ')
        del tribunal[0:2]
        tribunal = ' '.join(tribunal)
        fechaIngreso = web.find_element(By.XPATH, '/html/body/form/table[3]/tbody/tr[1]/td[3]').text.split(' ')[-1]
        estado = web.find_element(By.XPATH, '/html/body/form/table[3]/tbody/tr[3]/td[2]').text.split(' ')[-1]
        page.append((
          rol,
          tribunal,
          fechaIngreso,
          estado,
          link
        ))
        workBook.save(filename='resumen.xlsx')
    except:
      retryDict[tribunal_code] = f'Error en el link {link}'
  web.close()
  web.switch_to.window(web.window_handles[0])


web.quit()

with open('retry.json', 'w') as fp:
    json.dump(retryDict, fp)


msg = MIMEMultipart()
msg['From'] = '' # Customize
msg['To'] = '' # Customize
msg['Date'] = formatdate(localtime=True)
msg['Subject'] = f'Resumen causas {date.today().day - 1}/{date.today().month}/{date.today().year}'

msg.attach(MIMEText('Adjuntado resumen.'))

for f in ['resumen.xlsx', 'retry.json']:
    with open(f, "rb") as fil:
        part = MIMEApplication(
            fil.read(),
            Name=basename(f)
        )
    part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
    msg.attach(part)


smtp = smtplib.SMTP('smtp.gmail.com', 587) # Customize 
smtp.login('user@email.com', 'password') # Customize
smtp.sendmail('', '', msg.as_string()) # Customize
smtp.close()

if exists('resumen.xlsx'): remove('resumen.xlsx')
if exists('retry.json'): remove('retry.json')
