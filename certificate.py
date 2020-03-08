import xlrd
import smtplib
import getpass
import imghdr
from email.message import EmailMessage


path = "ncc.xlsx"
inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)

bcc = []
for i in range(inputWorksheet.nrows):
    bcc.append(inputWorksheet.cell_value(i, 3))

sender = 'credenzusser@gmail.com'
password = getpass.getpass("Enter your password:- ")

msg = EmailMessage()
msg['From'] = sender
msg['Subject'] = "This is just a test"
msg['Bcc'] = ', '.join(bcc)
msg.set_content(
    "Here is an attachment with the mail.\nPlease do not discose the certificate elsewhere.")

with open('certificate.jpeg', 'rb') as image:
    file_data = image.read()
    file_type = imghdr.what(image.name)
    file_name = image.name

msg.add_attachment(file_data, maintype='image', subtype=file_type)

with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login(sender, password)
    smtp.send_message(msg)
    print("Sent mails")
