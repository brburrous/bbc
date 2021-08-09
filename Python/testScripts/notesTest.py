import os
import xlrd
from progress.bar import FillingSquaresBar
import pyperclip
from icecream import ic



from tkinter import Tk, filedialog
root = Tk() # pointing root to Tk() to use it as Tk() in program.
root.withdraw() # Hides small tkinter window.

root.attributes('-topmost', True) # Opened windows will be active. above all windows despite of selection.

directory = filedialog.askdirectory(initialdir='/Users/brian/Developer') # Returns opened path as str
print(directory)





files = list()

for entry in os.scandir(directory):
    if entry.path.endswith('.xls') and entry.is_file():
        files.append(entry)
bar = FillingSquaresBar('Countdown', max = len(files))







def IsANote(sheet, row):
    lawSlice = sheet.row_slice(row, 4, 7) 
    costSlice = sheet.row_slice(row, 10, 14)
    slice = lawSlice + costSlice
    tot = 0
    for cell in slice:
        if cell.value != 0:
            tot += cell.ctype        
    return tot == 0





notesList= list()
codesSet = set()

for file in files:
    try: 
        book = xlrd.open_workbook(file)
        sheet = book.sheet_by_index(1)
  
        if sheet.cell(17, 3).value == 'CARPETS AND RUGS' and sheet.ncols == 14:
            i = 18
            # i = 18 # Row index of first room
            # while(sheet.cell(i, 3).value != 'UPHOLSTERY' and i < sheet.nrows):
            #     if sheet.cell(i, 3).ctype != 0 and sheet.cell(i,3).value.strip():
            #         if IsANote(sheet, i):
            #             notesList.append(sheet.cell(i, 3).value)
            #             codes = [i.ctype for i in sheet.row_slice(i, 7, 10)]
            #             if 1 in codes:
            #                 print(sheet.cell(i,3).value)
            #                 print(file.name)
                            
            #                 input("Press Enter to continue...")

            i = i + 1
 
    except:
        print(file)
        book = xlrd.open_workbook(file)
        sheet = book.sheet_by_index(1)
        print(sheet.nrows)
    #BAR.NEXT()
 
# print("Number of strings matching pattern: {}".format(len(notesList)))
# print(codesSet)
#bar.finish()

/
ar