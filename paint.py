from PIL import Image, ImageDraw, ImageFont
import xlrd

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


image = Image.open('certificate1.jpeg')
font = ImageFont.truetype(
    '/usr/share/fonts/truetype/freefont/FreeMono.ttf', size=25)

draw = ImageDraw.Draw(image)
draw.text(xy=(525, 430), text=user1[0], fill=(0, 0, 0), font=font)
image.save(f'{user1[0]}.png')
draw.text(xy=(525, 430), text=user2[0], fill=(0, 0, 0), font=font)
image.save(f'{user2[0]}.png')
print("Done")
