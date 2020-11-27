from PIL import Image, ImageDraw, ImageFont
import xlrd
import os

path = "participation.xlsx"
inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)
rows = inputWorksheet.nrows

user = []
objects = {}
for i in range(rows):
    objects['name'] = inputWorksheet.cell_value(i, 1)
    objects['team'] = inputWorksheet.cell_value(i, 0)
    user.append(objects)
    objects = {}


for person in user:
    image = Image.open('participation.jpg')
    name = person['name']
    team = person['team']
    font = ImageFont.truetype(
        f'{os.getcwd()}/Work.ttf', size=40)

    draw = ImageDraw.Draw(image)
    draw.text(xy=(630, 518), text=name, fill=(0, 0, 0), font=font)
    draw.text(xy=(590, 578), text=team, fill=(0,0,0), font=font)
    image.save(f'{os.getcwd()}/certificates/{team}_{name}.jpg')
    print(f'Generated for {team}_{name}')
