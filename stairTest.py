from Python.helperFuncs import getDirectory, getFiles
from Python/roomType.py import *
import os
import xlrd
from Python/helperFuncs import *

directory = getDirectory()

files = getFiles(directory)

for file in files:
    book = xlrd.open_workbook(file)
    sheet = book.sheet_by_index(1)
    