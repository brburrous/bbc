import os
import xlrd
import re
from progress.bar import FillingSquaresBar

from tkinter import Tk, filedialog
root = Tk() # pointing root to Tk() to use it as Tk() in program.
root.withdraw() # Hides small tkinter window.

root.attributes('-topmost', True) # Opened windows will be active. above all windows despite of selection.

directory = filedialog.askdirectory() # Returns opened path as str
print(directory)

files = list()

for entry in os.scandir(directory):
    if entry.path.endswith('.xls') and entry.is_file():
        files.append(entry)
bar = FillingSquaresBar('Countdown', max = len(files))

p = re.compile(r'(^\s*"\s+"|^\s*")')
quoteList = list()

for file in files: 
    book = xlrd.open_workbook(file)
    sheet = book.sheet_by_index(1)

    if sheet.cell(17, 3).value == 'CARPETS AND RUGS' and sheet.ncols == 14:

        i = 18 # Row index of first room
        while(sheet.cell(i, 3).value != 'UPHOLSTERY' and i < sheet.nrows):
            if p.match(sheet.cell(i,3).value):
                print(sheet.cell(i,3).value)
                quoteList.append(sheet.cell(i, 3).value)
            i = i + 1
    bar.next()
    
print("Number of strings matching pattern: {}".format(len(quoteList)))

bar.finish()