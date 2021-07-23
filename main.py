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

