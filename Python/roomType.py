
def IsANote(sheet, row):
    lawSlice = sheet.row_slice(row, 4, 7) 
    costSlice = sheet.row_slice(row, 10, 14)
    slice = lawSlice + costSlice
    tot = 0
    for cell in slice:
        if cell.value != 0:
            tot += cell.ctype        
    return tot == 0

def isStairs(sheet, row):
        lwSlice = sheet.row_slice(row, 4, 6)
        tot = 0
        for cell in slice:
            if cell.value != 0:
                tot += cell.ctype
        if tot != 0:
            return False    
        return

import re

def inconsistentStairs():
        p = re.compile('.*[sS][tT]')
        cCodecell = sheet.cell(row, 7)
        if p.match(cCodeCell):
