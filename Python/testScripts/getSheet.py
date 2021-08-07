import xlrd
from tkinter import Tk, filedialog
root = Tk() # pointing root to Tk() to use it as Tk() in program.
root.withdraw() # Hides small tkinter window.

root.attributes('-topmost', True) # Opened windows will be active. above all windows despite of selection.

file = filedialog.askopenfile()

book = xlrd.open_workbook(file.name)

sheet = book.sheet_by_index(1)


