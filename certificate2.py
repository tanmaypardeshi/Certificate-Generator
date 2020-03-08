import xlrd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

sender = 'credenzusser@gmail.com'
receiver = 'kaustubhodak1@gmail.com'
subject = 'This is just a test'

msg = MIMEMultipart()
msg['From'] = sender
msg['To'] = receiver
msg['Subject'] = subject

body = 'Here is an attachment with the mail.\nPlease do not discose the certificate elsewhere.'
msg.attach(MIMEText(body, 'plain'))

filename = 'certificate.jpeg'
attachment = open(filename, 'rb')

part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachement; filename= '+filename)

msg.attach(part)
text = msg.as_string()
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(sender, "credenztechdays")

server.sendmail(sender, receiver, text)
server.quit()
