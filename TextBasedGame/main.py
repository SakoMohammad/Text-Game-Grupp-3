from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

import re



workbook = load_workbook("Platsnamn textspel.xlsx")
sheet = workbook['Rum']
GameMap = workbook['Karta']
wsSave = workbook['Sparande']

room_description= 1




#sheet.append(["ID",'Room','Room Description', 'Paths'])



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
     
    print(str(room_value[room_description]))

def save(x,y):#- här kommer save funktionen

    wsSave['A1'] = x
    workbook.save("Platsnamn textspel.xlsx")
    print('Your game progress has now been saved')

def move(x,y,direction,valid_directions):#- här kommer rörelse funktionen

    direction.lower()
    if direction =='north'and direction in valid_directions:
        y-=1
    elif direction =='south' and direction in valid_directions:
        y+=1
    elif direction =='east' and direction in valid_directions:
        x+=1
    elif direction =='west' and direction in valid_directions:
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
        possible_directions= current_room_type[2].split(" ")


        a = input()
        if a == 'look':
           look(current_room_type)


        if a == 'save':
            save(x,y)

        if re.search('move ',a):
            text = a
            text = re.sub('move ', '', text)



            x,y = move(x,y,text,possible_directions)
            print(text)
        if a == 'quit':
           running= False


        player_position = main_map[y][x]
        current_room_type = room_IDs[int(float(player_position))]
        #print(player_position)

        #print(str(x) + ' ' + str(y))
        #print(str(current_room_type))





main()
#sheet.append(["1",'hallway','you see a brightly lit hallway infront of you', 'North,South'])

#workbook.save("hello_world.xlsx")

