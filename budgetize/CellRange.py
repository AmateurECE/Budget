###############################################################################
# NAME:             CellRange.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      The LibreOffice API doesn't have particularly great support
#                   for iterating over cell ranges in a very Python-ic way, so
#                   I wrote this glue logic to make it a little nicer on the
#                   eyes.
#
# CREATED:          12/01/2021
#
# LAST EDITED:      12/02/2021
###

import re

CELLRANGE_RE = re.compile(r'([A-Z]{1,3})([0-9]{1,7})')
COLUMN_MAX = 1023
ROW_MAX = 1048575

RowIteratorFn = lambda index, dim, access: access.getCellByPosition(dim, index)
ColumnIteratorFn = lambda index, dim, access: access.getCellByPosition(
    index, dim)

class CellListIterator:
    def __init__(self, row, columns, xIndexAccess, iteratorFn):
        self.row = row
        self.columns = columns
        self.index = 0
        self.xIndexAccess = xIndexAccess
        self.iteratorFn = iteratorFn

    def __next__(self):
        if self.index >= self.columns:
            raise StopIteration()
        index = self.index
        self.index += 1
        return self.iteratorFn(index, self.row, self.xIndexAccess)

class CellMatrixIterator:
    def __init__(self, rows, columns, xIndexAccess):
        self.rows = rows
        self.columns = columns
        self.index = 0
        self.xIndexAccess = xIndexAccess

    def __next__(self):
        if self.index >= self.rows:
            raise StopIteration()
        index = self.index
        self.index += 1
        return CellRangeContainer(lambda: CellListIterator(
            index, self.columns, self.xIndexAccess, RowIteratorFn))

class CellRangeContainer:
    def __init__(self, iteratorFn):
        self.iteratorFn = iteratorFn

    def __iter__(self):
        return self.iteratorFn()

def getColumnFromString(rowString):
    rowBytes = bytearray(rowString, 'ascii')
    index = len(rowBytes) - 1
    value = 0
    for byte in rowBytes:
        value += (byte - 64) * (26**index)
        index -= 1
    return value - 1

def getCoordinatesForCellSpec(spec):
    match = re.match(CELLRANGE_RE, spec)
    if not match:
        return None

    column = getColumnFromString(match.group(1))
    row = int(match.group(2)) - 1
    if column > COLUMN_MAX or row > ROW_MAX:
        return None
    return (column, row)

def getCellRangeForMatrixSpec(spec, xSheet):
    firstSpec, secondSpec = tuple(spec.split(':'))
    firstRow, firstColumn = getCoordinatesForCellSpec(firstSpec)
    secondRow, secondColumn = getCoordinatesForCellSpec(secondSpec)
    xIndexAccess = xSheet.getCellRangeByName(spec)
    if firstRow == secondRow:
        return CellRangeContainer(lambda: CellListIterator(
            0, secondColumn - firstColumn, xIndexAccess, RowIteratorFn))
    if firstColumn == secondColumn:
        return CellRangeContainer(lambda: CellListIterator(
            0, secondRow - firstRow, xIndexAccess, ColumnIteratorFn))
    return CellRangeContainer(lambda: CellMatrixIterator(
        secondRow - firstRow, secondColumn - firstColumn, xIndexAccess))

###############################################################################
