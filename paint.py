from PIL import Image, ImageDraw, ImageFont
import xlrd

path = "NCC.xlsx"
inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)

user1 = []
user2 = []
bcc = []
for i in range(inputWorksheet.nrows):
    user1.append(inputWorksheet.cell_value(i, 1))
    user2.append(inputWorksheet.cell_value(i, 2))
    bcc.append(inputWorksheet.cell_value(i, 3))

for i in range(inputWorksheet.nrows):
    image1 = Image.open('certificate.jpeg')
    image2 = Image.open('certificate.jpeg')

    font = ImageFont.truetype(
        '/usr/share/fonts/truetype/freefont/FreeMono.ttf', size=25)
    draw = ImageDraw.Draw(image1)
    draw.text(xy=(545, 435), text=user1[i], fill=(0, 0, 0), font=font)
    image1.save(f'{user1[i]}.png')
    print(f'Generated certificate {i+1} for:-\n{user1[i]}')
    draw = ImageDraw.Draw(image2)
    draw.text(xy=(545, 435), text=user2[i], fill=(0, 0, 0), font=font)
    if user2[i] == '':
        pass
    else:
        print(f'{user2[i]}')
        image2.save(f'{user2[i]}.png')
