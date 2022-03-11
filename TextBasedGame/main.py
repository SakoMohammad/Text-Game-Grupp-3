from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

import re



workbook = load_workbook("Platsnamn_textspel.xlsx")
sheet = workbook['Rum']
GameMap = workbook['Karta']






#sheet.append(["ID",'Room','Room Description', 'Paths'])



workbook.create_sheet('PlayerInfo')


mapx=0
mapy=0

def goThroughSheet(thesheet):
    # goes through a sheet and puts every value into a 2d matrix (I think i can use that word)
    main_map=[] # the matrix that is going to be returned
    #row_number=0
    #column_number=0
    for row in thesheet:
        #row_number+=1
        column_rooms=[]
        for column in row:
            #column_number+=1
            column_rooms.append(str(column.value))#adds a cell from sheet row to the lists row
        main_map.append(column_rooms)
    return main_map

def look(room_value):
     
    print(str(sheet.cell(room_description,1).value))

def save(direction):# här kommer spar funktionen
    pass

def move(x,y,direction):#- här kommer rörelse funktionen

    direction.lower()
    if direction =='north':
        y-=1
    elif direction =='south':
        y+=1
    elif direction =='east':
        x+=1
    elif direction =='west':
        x-=1
    else:
        print('Enter a valid direction')
    return x , y


#print(goThroughSheet(GameMap))

def main():
    x = 0
    y = 1
    room_description = 3

    running = True
    room_IDs = goThroughSheet(sheet)
    main_map = goThroughSheet(GameMap)
    player_position = main_map[y][x]
    current_room_type = room_IDs[int(float(player_position))]

    while running:



        a = input()
        if a == 'look':
           look()


        if a == 'save':
            save()

        if re.search('move ',a):
            text = a
            text = re.sub('move ', '', text)



            x,y = move(x,y,text)
            print(text)
        if a == 'quit':
           running= False


        player_position = main_map[y][x]
        current_room_type = room_IDs[int(float(player_position))]
        print(player_position)

        print(str(x) + ' ' + str(y))
        print(str(current_room_type))





main()
#sheet.append(["1",'hallway','you see a brightly lit hallway infront of you', 'North,South'])

#workbook.save("hello_world.xlsx")

