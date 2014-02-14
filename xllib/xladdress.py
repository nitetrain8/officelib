"""
Created on Jan 29, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather


Hold functions related to creating/converting
xl based cell address to and from python indicies.

"""


def cellStr(row, col, AbsoluteColumn=False, AbsoluteRow=False):

    """ Convert row, column index rep into A1 style rep.
        if AbsoluteColumn or AbsoluteRow
        are True or 0, give abs notation for that row
        or column ie $A$1.

        Finally got string index thing working. 1-based index made
        it really confusing.

        @param row: the row number eg 1 -> row 1
        @type row: int
        @param col: the column number eg 1 -> column A
        @type col: int
        @param AbsoluteColumn: use absolute column (insert '$' before column)
        @type AbsoluteColumn: int
        @param AbsoluteRow: use absolute row (insert '$' before row)
        @type AbsoluteRow: int
        @return: the excel-style spreadsheet range address.
        @rtype: str
    """

    if row < 1 or col < 1:
        raise ValueError("Values for 'row' and 'col' must be greater than 0.")
        
    target = ''

    while col > 26:
        target = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[(col - 1) % 26] + target
        col = (col - 1) // 26

    AbsoluteColumn = '$' if AbsoluteColumn else ''
    AbsoluteRow = '$' if AbsoluteRow else ''

    return ''.join((
                    AbsoluteColumn, 
                    'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[col - 1], 
                    target, 
                    AbsoluteRow, 
                    str(row)))


def cellRangeStr(cell1, cell2, cellStr=cellStr):
    """ For those times you need an $A$1:$B$2 style
        cell reference

    @param cell1: each is a tuple of (row/column[abscol[absrow]])
                    automatically unpacked by python and sent to cellStr.
                    See cellStr doc for params to cellStr
    @type cell1: (int, int) | (int, int, int) | (int, int, int, int)
    @type cell2: (int, int) | (int, int, int) | (int, int, int, int)
    @param cell2: same as cell1

    @return: cellStr addresses joined by ':' eg A$1:$B$2
    @rtype: str
    """
        
    return ':'.join((cellStr(*cell1), cellStr(*cell2)))


def xlrange(start, stop, step=1):
    """excel uses 1-based index, use xlrange
    to convert 0-index based range generator to
    excel's range format.
    """
    return range(start, stop + step, step)
        

# Export Aliases
cellRange = cellRangeStr
cellAddress = cellStr


if __name__ == '__main__':
    pass




