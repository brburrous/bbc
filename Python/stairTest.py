from roomType import *
import os
import xlrd
from helperFuncs import *
import pyperclip


directory = getDirectory()

files = getFiles(directory)

for file in files:
    book = xlrd.open_workbook(file)
    sheet = book.sheet_by_index(1)

    i = 18
    while i < sheet.nrows and sheet.cell(i, 3).value != 'UPHOLSTERY':
        if inconsistentStairs(sheet, i):
            print("Inconsistent Stairs: {}".format(sheet.cell(i, 3).value))
            input("Press ENTER to continue...")
            pyperclip.copy(file.name)
    