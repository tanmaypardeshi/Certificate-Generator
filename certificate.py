import xlrd
import smtplib


path = "ncc.xlsx"
inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)

userlist = []

for i in range(inputWorksheet.nrows):
	dictionary = {
		"username1": inputWorksheet.cell_value(i, 1),
		"username2": inputWorksheet.cell_value(i, 2),
		"emailid": inputWorksheet.cell_value(i, 3),

	}
	userlist.append(dictionary)

print(userlist)

with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
     smtp.ehlo()
     smtp.starttls()
     smtp.ehlo()
     
     sender = "credenzusser@gmail.com"
     receiver = "kajalparekh@gmail.com"
     password = input("Enter your password:- ")
     
     smtp.login(sender,password)
     
     subject = "Regarding NCC"
     body = "Thank you participating in NCC 2020. \nHere is your certificate"
     
     msg = f'Subject: {subject}\n\n{body}'
     
     smtp.sendmail(sender, receiver, msg)
     


