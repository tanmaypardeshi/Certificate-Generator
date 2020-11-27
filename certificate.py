import xlrd
import os
import smtplib
import getpass
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

sender = 'pisb.credenz20@gmail.com'
password = getpass.getpass('Enter password:- ')
subject = 'PISB Ideathon 2020 Participation Certificate'

path = "participation.xlsx"
inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)
rows = inputWorksheet.nrows

user = []
objects = {}

for i in range(rows):
    objects['email'] = inputWorksheet.cell_value(i,2)
    objects['name'] = inputWorksheet.cell_value(i, 1)
    objects['team'] = inputWorksheet.cell_value(i, 0)
    user.append(objects)
    objects = {}

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(sender, password)

for person in user:
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['Subject'] = subject

        team = person['team']
        name = person['name']
        email = person['email']

        body = f'Hi {name}, \n\nThank you participating in the PISB Ideathon 2020.\nWe hope you enjoyed it and would love to see you in future events by PISB, Pune.' \
                   f'\n\nIt would be wonderful if you share your certificate on social media with hashtag #pisb #ideathon\n\nThanks and Regards, PICT IEEE Student Branch'
        msg.attach(MIMEText(body, 'plain'))

        file = f'{os.getcwd()}/certificates/{team}_{name}.jpg'
        attachment = open(file, 'rb')
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachement; filename= ' + f'{team}_{name}.jpg')
        msg.attach(part)
        text = msg.as_string()
        server.sendmail(sender, email, text)
        print(f'Sent mail to {email}')
    except Exception as e:
        print(e.__str__(), f'Could not send an email to {email}')
server.quit()
