import xlrd

from getDirectory import * 
directory = getDirectory()
files = getFiles(directory)

namesFile = open('names.txt', 'w')
titlesFile = open('titles.txt', 'w')

numBadName = 0
numBadTitle = 0

    

for file in files: 
    book = xlrd.open_workbook(file)
    sheet = book.sheet_by_index(1)
    if sheet.cell(16, 3).value != "Description of Services":
        print("Filename of sheet with different name: {}".format(file))
        numBadName = numBadName+1
    if sheet.cell(17, 3).value != "CARPETS AND RUGS":
        print("Filename of sheet with different title: {}".format(file))
        numBadTitle = numBadTitle +1
    


print( "Number of bad names: {}".format(numBadName))
print( "Number of bad titles: {}".format(numBadTitle))







