from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter



workbook = load_workbook("hello_world.xlsx")
sheet = workbook['Sheet']
GameMap = workbook['GameMap']
RoomData = workbook['RoomData']
library=workbook.active


x=0
y=0


#sheet.append(["ID",'Room','Room Description', 'Paths'])
print(library.cell(1,1))

workbook.create_sheet('PlayerInfo')

library.cell(1,1)
mapx=0
mapy=0


for row in range(1,3):
    for column in range(1,5):
        print(sheet.cell(row,column).value)


def look(roomx,roomy):

    print(sheet.cell(roomx,roomy).value)

look(3,1)


for worksheet in workbook:
    print(worksheet)





#workbook.save("hello_world.xlsx")