import xlrd
import os

from tkinter import Tk, filedialog
root = Tk() # pointing root to Tk() to use it as Tk() in program.
root.withdraw() # Hides small tkinter window.

root.attributes('-topmost', True) # Opened windows will be active. above all windows despite of selection.

directory = filedialog.askdirectory() # Returns opened path as str
print(directory)

files = list()

for entry in os.scandir(directory):
    if (entry.path.endswith('.xls') 
            or entry.path.endswith('.xlsx')) and entry.is_file():
        files.append(entry)

for file in files:
    print(file.name)

book = xlrd.open_workbook(files[0].path)


print("Worksheet name(s): {0}".format(book.sheet_names()))

sheet = book.sheet_by_index(0)

row = sheet.row_slice(1)
newRow = list(filter(lambda x: x.ctype != 0, row))
finalRow = [x.value for x in newRow]

book = xlrd.open_workbook("title", formatting_info=True)
sheet = book.sheet_by_index(1)
cell = sheet.cell(17,5)
xfx = cell.xf_index
xf = book.xf_list[xfx]
bg = xf.background.pattern_colour_index

colorValues = {yellow:26, white:64, grey:22}

def getColor(cell):
    xfx = cell.xf.index
    xf  = book.xf_list[xfx]
    bg = xf.background_colour_index


# class jobCustomer {
#     name: string,
#     email: string,
#     homeNum: number,
#     workNum: number,
#     cellNum: number,
#     otherNum: number;
#     index = (name: 'B5', email:'B10', homeNum:'B12', workNum: 'B13', cellNum: 'B14', otherNum: 'B15')
# }

# class billCustomer {
#     name: string,
#     email: string,
#     homeNum: number,
#     workNum: number,
#     cellNum: number,
#     otherNum: number
#     index = (name: 'B5', email:'F10', homeNum:'F12', workNum: 'F13', cellNum: 'F14', otherNum: 'F15')
# }

# class address {
#     zip ,
#     streetName ,
#     streetNum , 
#     suffix ,
#     name ,
#     filename , 
#     refNum 
#     index = (zip: 'B21', streetName:'B22', streetNum:'B23', suffix: 'B24', name: 'B25', filename: 'B27', refNum: 'B29')

# }

# class room {
#     service; 
#     length; 
#     width; 
#     cleaningCode; 
#     protectionCode; 
#     furnitureCode; 


# }


# jobCustomer.name:'B5'

def extractData():
    print("Hello world")

if (sheet.ncols == 14):
    extractData()

def printRooms():
    if sheet.cell(17, 3) == 'CARPETS AND RUGS':
        i = 18
        while(sheet.cell(i, 3) != 'UPHOLSTERY' and i < sheet.nrows):
            if sheet.cell(i, 3).ctype != 0:
                print(sheet.cell.value)

print("Hello World")




# A comment