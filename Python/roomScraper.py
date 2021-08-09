# data/94920teal0025.xls
import xlrd
from roomClass import room
import re

p = re.compile(r'(^\s*"\s+"|^\s*")')

def extractRoomsData(sheet):
    # Check to ensure sheet is properly structured
    if sheet.cell(17, 3).value == 'CARPETS AND RUGS':
        roomsList = list() # Create new list to store all rooms

        i = 18 # Row index of first room
        while(sheet.cell(i, 3).value != 'UPHOLSTERY' and i < sheet.nrows):
            # If cell contains an alphanumeric string, it is a room
            # This data from this room is extracted
            if sheet.cell(i, 3).ctype != 0 and any(c.isalnum() for c in sheet.cell(i,3).value):
                print("Room #{} with name {}".format((i-17), sheet.cell(i,3).value))
                roomInstance = getRoomData(sheet, i)

                if p.match(roomInstance.name):
                    roomInstance.name = p.sub(roomsList[-1].name, roomInstance.name)
                # Add roomInstance object of type 'room'
                roomsList.append(roomInstance)
            i = i + 1

    else:
        print("error")
        print(sheet.cell(17,3))

    return roomsList


def cellType(sheet, row):
    # check if there is a number in the length column
    if (sheet.cell(row, 4).ctype == 2): 
        # do a thing
        print("hello")
    # check if there is a moving fee
    elif (sheet.cell(row, 11) == 2):
        # do another thing
        print("Hello")
    # If other conditions fail, cell is a note
    else:
        print("hello")


def getRoomData(sheet, row):
    # Extract Revevant room properties and store in an instance of 
    # class 'room' 
    roomInstance = room(sheet.cell(row, 3).value)
    roomInstance.length = sheet.cell(row, 4).value
    roomInstance.width = sheet.cell(row, 5).value
    roomInstance.cleanCode = sheet.cell(row, 7).value
    roomInstance.furnitureCode = sheet.cell(row, 8).value
    roomInstance.protectionCode = sheet.cell(row, 9).value

    return roomInstance




'num1',
'num2',
'This is a test',
'num4',
'num5',
'num6',
'num7',
'num8',
'num9',
'num10',
'num1',
'num2',
'num3',
'num4',