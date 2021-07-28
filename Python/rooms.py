# data/94920teal0025.xls
import xlrd
import os

filename = 'data/94920teal0025.xls' 


def printRooms(sheet):
    if sheet.cell(17, 3).value == 'CARPETS AND RUGS':
        i = 18
        while(sheet.cell(i, 3).value != 'UPHOLSTERY' and i < sheet.nrows):
            if sheet.cell(i, 3).ctype != 0 and any(c.isalnum() for c in sheet.cell(i,3).value):
                print("Room #{} with name {}".format((i-17), sheet.cell(i,3).value))
            i = i + 1
    else:
        print("error")
        print(sheet.cell(17,3))

def callPrintRooms(sheetName):
    book = xlrd.open_workbook(sheetName)
    sheet = book.sheet_by_index(1)
    printRooms(sheet)

