import xlrd
import smtplib
import ssl


def getusers():
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
