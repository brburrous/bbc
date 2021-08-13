
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
	area = sheet.cell(row, 6)
	tot = 0
	for cell in lwSlice:
		if cell.value != 0:
			tot += cell.ctype
	if tot != 0:
		return False    
	return (area.ctype != 0 and area.value != 0)


def inconsistentStairs(sheet, row):
	if isStairs(sheet, row):
		import re
		p1 = re.compile(r".*[sS][tT]|.*[bB][bB]")
		cCodeCell = sheet.cell(row, 7).value
		if not p1.match(cCodeCell):
			return True
	return False
