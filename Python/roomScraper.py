# data/94920teal0025.xls
import xlrd
from roomClass import room

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

                # Extract Revevant room properties and store in an instance of 
                # class 'room' 
                roomInstance = room(sheet.cell(i, 3).value)
                roomInstance.length = sheet.cell(i, 4).value
                roomInstance.width = sheet.cell(i, 5).value
                roomInstance.cleanCode = sheet.cell(i, 7).value
                roomInstance.furnitureCode = sheet.cell(i, 8).value
                roomInstance.protectionCode = sheet.cell(i, 9).value

                # Add roomInstance object of type 'room'
                roomsList.append(roomInstance)

            i = i + 1

    else:
        print("error")
        print(sheet.cell(17,3))

    return roomsList






