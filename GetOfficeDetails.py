from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import openpyxl
import os
import re

import smtplib
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


driver = webdriver.Chrome()

driver.get('https://my.fibank.bg/EBank/public/offices')
driver.maximize_window()
wait = WebDriverWait(driver, 10)


root = wait.until(EC.presence_of_all_elements_located((By.XPATH, 
                                                       "//div[contains(@class, 'col col-full')][.//ul[contains(@class, 'list-inline')]]")))
officeElements = root[0].find_elements(By.XPATH,
                             "//div[contains(@class, 'margin-16')]")

pathToExcel = 'C:\\PythonApp\\fibank_branches.xlsx'

if os.path.exists(pathToExcel):
    os.remove(pathToExcel)
    print('Existing file "{pathToExcel}" deleted.')

wb = openpyxl.Workbook()
ws = wb.active

holidayPattern = r'(?:Събота|Неделя)'

idx = 0

for i, item in enumerate(officeElements, start=1):
    if re.search(holidayPattern, item.text):
        idx+= 1
        ws.cell(row=idx, column=1, value=item.text)

wb.save(pathToExcel)
print('File Created Successfully')

sender_email = 'encho871337@gmail.com'
recipients = ['encho871337@gmail.com']
subject = 'Subject: Sending Excel File via Gmail SMTP'
body = 'This is the body of the email'
message = MIMEMultipart()
message.attach(MIMEBase('application', 'octet-stream'))
message['From'] = sender_email
message['To'] = ', '.join(recipients)
message['Subject'] = subject
message.attach(MIMEBase('application', 'octet-stream'))
part = MIMEBase('application', 'octet-stream')
part.set_payload(open(f'{pathToExcel}', 'rb').read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename="fibank_branches.xlsx"')
message.attach(part)

with smtplib.SMTP('smtp.gmail.com', 587) as server:
    server.starttls()  
    server.login(sender_email, 'nelv wffd wiie fnrg') 

    # Send the email
    server.sendmail(sender_email, recipients, message.as_string())

print('Email sent successfully!')

