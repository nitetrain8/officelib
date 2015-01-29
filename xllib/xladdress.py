"""
Created on Jan 29, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather


Hold functions related to creating/converting
xl based cell address to and from python indicies.

"""
from officelib import const


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


def chart_range_strs(xcol, ycol, top, bottom, ws_name=''):
    """
    @param xcol: x column
    @param ycol: y column
    @param top: top row
    @param bottom: end row
    @param ws_name: name of worksheet
    @return: (str, str)

    One of the most common uses of cellRangeStr is to add a set of
    columns as the source data for a chart. This function makes that
    easier.
    """
    # ws_name must be quoted + exclamation in excel formula
    if ws_name:
        ws_name = "='%s'!" % ws_name
    else:
        ws_name = "="

    xrng = cellRangeStr(
        (top, xcol), (bottom, xcol)
    )

    yrng = cellRangeStr(
        (top, ycol), (bottom, ycol)
    )

    xrng = ws_name + xrng
    yrng = ws_name + yrng

    return xrng, yrng


def xlrange(start, stop, step=1):
    """excel uses 1-based index, use xlrange
    to convert 0-index based range generator to
    excel's range format.
    """
    return range(start, stop + step, step)
        

def column_pair_range_str_by_header(cells, header):
    cell = cells.Find(What=header, After=cells(1, 1), SearchOrder=const.xlByRows)
    xcol = cell.Column
    ycol = xcol + 1
    top = cell.Row + 1
    bottom = cell.End(const.xlDown).Row
    return chart_range_strs(xcol, ycol, top, bottom, cells.Worksheet.Name)

# Export Aliases
cellRange = cellRangeStr
cellAddress = cellStr


if __name__ == '__main__':
    pass




