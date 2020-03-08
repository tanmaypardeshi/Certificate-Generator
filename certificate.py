import xlrd
import smtplib
import getpass


path = "ncc.xlsx"
inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)


with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
    smtp.ehlo()
    smtp.starttls()
    smtp.ehlo()
    sender = "credenzusser@gmail.com"
    password = getpass.getpass("Enter your password:- ")

    smtp.login(sender, password)

    for i in range(inputWorksheet.nrows):
        receiver = inputWorksheet.cell_value(i, 3)
        user1 = inputWorksheet.cell_value(i, 1)
        user2 = inputWorksheet.cell_value(i, 2)
        subject = "Regarding NCC"
        if user2 == '':
            body = f'Thank you {user1} for participating in NCC 2020'
        else:
            body = f'Thank you {user1} and {user2} for participating in NCC 2020.'
        msg = f'Subject: {subject}\n\n{body}'

        smtp.sendmail(sender, receiver, msg)
        print(f'Sent mail {i+1}')
