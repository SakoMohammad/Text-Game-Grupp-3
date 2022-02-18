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
room_description= 3

running = True





print(str(library.cell(1,1).value))

workbook.create_sheet('PlayerInfo')

library.cell(1,1)
mapx=0
mapy=0

def goThroughSheet(thesheet): # the function that will load in the game map into the game
    main_map=[]
    row_number=0
    column_number=0
    for row in sheet:
        row_number+=1
        column_rooms=[]
        for column in row:
            column_number+=1
            column_rooms.append(str(thesheet.cell(row_number, column_number).value))
        main_map.append(column_rooms)
    return main_map


def look(room_value):
     
    print(str(sheet.cell(room_description, 1).value))


def save(direction):  # här kommer spar funktionen
    pass


def move():  # här kommer rörelse funktionen
    pass


print(str(goThroughSheet(sheet)))

while running:

    a = input()
    if a == 'look':
        print('works')
        if str(library.cell(1,1).value) == "ID":
            print('can string')


    if a == 'save':
        save()

    if re.search('move ',a):
        text = a
        text = re.sub('move ', '', text)

        print (text)
        move(text)

    if a == 'quit':
       running= False





#sheet.append(["1",'hallway','you see a brightly lit hallway infront of you', 'North,South'])

#workbook.save("hello_world.xlsx")

