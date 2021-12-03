###############################################################################
# NAME:             cellname.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Logic for handling Cell names
#
# CREATED:          12/02/2021
#
# LAST EDITED:      12/02/2021
###

import re

CELLRANGE_RE = re.compile(r'([A-Z]{1,3})([0-9]{1,7})')
COLUMN_MAX = 1023
ROW_MAX = 1048575

def getIndexFromColumnName(rowString):
    rowBytes = bytearray(rowString, 'ascii')
    index = len(rowBytes) - 1
    value = 0
    for byte in rowBytes:
        value += (byte - 64) * (26**index)
        index -= 1
    return value - 1

def getColumnNameFromIndex(index):
    rowBytes = []
    index += 1
    while index > 0:
        rem = index % 26
        index //= 26
        rowBytes.insert(0, rem + 64)
    return bytearray(rowBytes).decode('ascii')

def getCoordinatesFromCellName(spec):
    match = re.match(CELLRANGE_RE, spec)
    if not match:
        return None

    column = getIndexFromColumnName(match.group(1))
    row = int(match.group(2)) - 1
    if column > COLUMN_MAX or row > ROW_MAX:
        return None
    return (column, row)

def getCellNameFromCoordinates(column, row):
    return f'{getColumnNameFromIndex(column)}{row + 1}'

###############################################################################
