from roomType import *
import os
import xlrd
from helperFuncs import *
import pyperclip
from progress.bar import FillingSquaresBar


directory = getDirectory()

files = getFiles(directory)
numExceptions = 0
bar = FillingSquaresBar('Countdown', max = len(files))

for file in files:
    try:    
        book = xlrd.open_workbook(file)
        sheet = book.sheet_by_index(1)

        if ensureFormat(sheet):
            i = 18
            while i < sheet.nrows and sheet.cell(i, 3).value != 'UPHOLSTERY':
                if inconsistentStairs(sheet, i):
                    numExceptions = numExceptions + 1
                    # print("Inconsistent Stairs: {}".format(sheet.cell(i, 3).value))
                    # pyperclip.copy(file.name)
                    # input("Press ENTER to continue...")
                i = i + 1
        else:
            print("File doesn't match format: {}".format(file.name))
            os.rename(file.path, '/Users/brian/Developer/BBC Residences/Format Issues/'+file.name)
    except:
        print("File threw an error: {}".format(file.name))
        os.rename(file.path, '/Users/brian/Developer/BBC Residences/Errors Thrown/'+file.name)
    bar.next()

bar.finish()
print(numExceptions)
    