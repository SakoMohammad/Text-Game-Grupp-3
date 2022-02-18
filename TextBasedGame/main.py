from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

import re



workbook = load_workbook("Platsnamn_textspel.xlsx")
sheet = workbook['Rum']
GameMap = workbook['Karta']




x=3
y=2
room_description= 3

running = True



#sheet.append(["ID",'Room','Room Description', 'Paths'])



workbook.create_sheet('PlayerInfo')


mapx=0
mapy=0

def goThroughSheet(thesheet):
    main_map=[]
    row_number=0
    column_number=0
    for row in thesheet:
        row_number+=1
        column_rooms=[]
        for column in row:
            column_number+=1
            column_rooms.append(str(thesheet.cell(row_number, column_number).value))
        main_map.append(column_rooms)
    return main_map

def look(room_value):
     
    print(str(sheet.cell(room_description,1).value))

def save(direction):# här kommer spar funktionen
    pass

def move():#- här kommer rörelse funktionen
    pass


for row in goThroughSheet(sheet):
    print(row)



while running:

    a = input()
    if a == 'look':
       look()


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

