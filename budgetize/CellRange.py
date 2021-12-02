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
# LAST EDITED:      12/01/2021
###

import re

CELLRANGE_RE = re.compile(r'([A-Z]{1,3})([0-9]{1,7})')
COLUMN_MAX = 1023
ROW_MAX = 1048575

class CellRangeIterator:
    def __init__(self, items, xIndexAccess, iterType):
        self.items = items
        self.xIndexAccess = xIndexAccess
        self.iterType = iterType
        self.index = 0

    def __next__(self):
        if self.index >= len(self.items):
            raise StopIteration()
        index = self.index
        self.index += 1
        if self.iterType == 'list':
            row, column = self.items[index]
            return self.xIndexAccess.getCellByPosition(row, column)
        return self.items[index]

class CellRangeContainer:
    def __init__(self, items, xIndexAccess, iterType):
        self.items = items
        self.xIndexAccess = xIndexAccess
        self.iterType = iterType

    def __iter__(self):
        return CellRangeIterator(self.items, self.xIndexAccess, self.iterType)

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
        return CellRangeContainer([(firstRow, i) for i in range(
            firstColumn, secondColumn + 1)], xIndexAccess, 'list')
    ranges = []
    for i in range(firstRow, secondRow + 1):
        ranges.append(CellRangeContainer([(j, i) for j in range(
            firstColumn, secondColumn + 1)], xIndexAccess, 'list'))
    return CellRangeContainer(ranges, None, 'matrix')

###############################################################################
