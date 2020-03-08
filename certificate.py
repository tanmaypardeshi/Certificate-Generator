import xlrd
import smtplib
import getpass
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

path = "ncc.xlsx"
inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)

user1 = []
user2 = []
receiver = []

for i in range(inputWorksheet.nrows):
    user1.append(inputWorksheet.cell_value(i, 1))
    user2.append(inputWorksheet.cell_value(i, 2))
    receiver.append(inputWorksheet.cell_value(i, 3))

sender = 'credenzusser@gmail.com'
password = getpass.getpass('Enter password:- ')
subject = 'This is just a test'


for j in range(inputWorksheet.nrows):
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['Subject'] = subject

    body = 'Here is an attachment with the mail.\nPlease do not discose the certificate elsewhere.'
    msg.attach(MIMEText(body, 'plain'))

    files = [f'{user1[j]}.png', f'{user2[j]}.png']
    for filename in files:
        if filename == '.png':
            pass
        else:
            attachment = open(filename, 'rb')
            part = MIMEBase('application', 'octet-stream')
            part.set_payload((attachment).read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',
                            'attachement; filename= '+filename)

            msg.attach(part)
            text = msg.as_string()
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(sender, password)
    server.sendmail(sender, receiver[j], text)
    print(f'Sent mail {j+1}')
    server.quit()
