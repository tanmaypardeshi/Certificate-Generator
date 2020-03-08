import xlrd
import smtplib
import getpass
import imghdr
from email.message import EmailMessage
from PIL import Image, ImageDraw, ImageFont


path = "ncc.xlsx"
inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)

user1 = []
user2 = []
bcc = []
for i in range(inputWorksheet.nrows):
    user1.append(inputWorksheet.cell_value(i, 1))
    user2.append(inputWorksheet.cell_value(i, 2))
    bcc.append(inputWorksheet.cell_value(i, 3))

sender = 'credenzusser@gmail.com'
password = getpass.getpass("Enter your password:- ")

msg = EmailMessage()
msg['From'] = sender
msg['Subject'] = "This is just a test"
msg['Bcc'] = ', '.join(bcc)
msg.set_content(
    "Here is an attachment with the mail.\nPlease do not discose the certificate elsewhere.")

PaintCertificate(user1, user2)

files = ['certificate1.jpeg', 'certificate2.jpeg']

for file in files:
    with open(file, 'rb') as image:
        file_data = image.read()
        file_type = imghdr.what(image.name)
        file_name = image.name

    msg.add_attachment(file_data, maintype='image',
                       subtype=file_type, filename=file_name)

with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login(sender, password)
    smtp.send_message(msg)
    print("Sent mails")


def PaintCertificate(user1, user2):

    image = Image.open('certificate1.jpeg')
    font = ImageFont.truetype(
        '/usr/share/fonts/truetype/freefont/FreeMono.ttf', size=25)

    draw = ImageDraw.Draw(image)
    draw.text(xy=(525, 430), text=user1[0], fill=(0, 0, 0), font=font)
    image.save(f'{user1[0]}.png')
    draw.text(xy=(525, 430), text=user2[0], fill=(0, 0, 0), font=font)
    image.save(f'{user2[0]}.png')
    print("Done")
