from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

import re



workbook = load_workbook("hello_world.xlsx")
sheet = workbook['Sheet']
GameMap = workbook['GameMap']
RoomData = workbook['RoomData']
library=workbook.active


x=3
y=2

running = True



#sheet.append(["ID",'Room','Room Description', 'Paths'])
print(library.cell(1,1))

workbook.create_sheet('PlayerInfo')

library.cell(1,1)
mapx=0
mapy=0

def goThroughSheet(thesheet):
    row_number=0
    column_number=0
    for row in sheet:
        row_number+=1
        for column in row:
            column_number+=1
            print(thesheet.cell(row_number,column_number).value)

def look(roomx,roomy):
     
    print(sheet.cell(3,1).value)

def save():# här kommer spar funktionen
    pass

def move():#- här kommer rörelse funktionen
    pass





while running:

    a = input()
    if a == 'look':
        print('works')

    if a == 'save':
        save()

    if re.search('move ',a):
        move(a)





#sheet.append(["1",'hallway','you see a brightly lit hallway infront of you', 'North,South'])

#workbook.save("hello_world.xlsx")

