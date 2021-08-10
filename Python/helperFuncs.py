
def getDirectory():
    from tkinter import Tk, filedialog
    root = Tk() # pointing root to Tk() to use it as Tk() in program.
    root.withdraw() # Hides small tkinter window.

    root.attributes('-topmost', True) # Opened windows will be active. above all windows despite of selection.

    directory = filedialog.askdirectory() # Returns opened path as str

    return directory

def getFiles(directory):
    import os
    files = list()
    for entry in os.scandir(directory):
        if (entry.path.endswith('.xls') 
                    or entry.path.endswith('.xlsx')) and entry.is_file():
                files.append(entry)
    return files

def getBlock(sheet, cellStart, cellEnd, offset=0):
    block = list()
    rowStart, colStart = str2tup(cellStart)
    rowFin, colFin = str2tup(cellEnd)
    for row in range(rowStart, rowFin+1):
        if row == rowStart:
            block = block + sheet.row_slice(row, colStart+offset, colFin+1)
        else:
            block = block + sheet.row_slice(row, colStart, colFin+1)
            
    block2 = [i.value for i in block]
    blockStr = ""
    blockStr = blockStr.join(block2)
    blockStr = blockStr.strip()
    return blockStr

# Convert excel cell index (ex: 'B12') to zero indexed tupple (row, col)
# str2tup('A12') -> (11, 0)
def  str2tup(cellIndex):
    let = cellIndex[0]
    col = 0
    if let.isupper():
        col = ord(let) - ord('A')
    else:
        col = ord(let) - ord('a')
    row = int(cellIndex[1:]) - 1
    return (row, col)


def ensureFormat(sheet):
    if cellV(sheet, 'D17') == 'Description of Services':
        return True
    return False


def cellV(sheet, cell):
    return sheet.cell(*str2tup(cell)).value