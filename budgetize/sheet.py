###############################################################################
# NAME:             Sheet.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Abstractions for handling sheets
#
# CREATED:          12/02/2021
#
# LAST EDITED:      02/06/2022
###

import uno
from com.sun.star.container import NoSuchElementException
from com.sun.star.sheet import CellFlags

from .cellrange import CellMatrix
from . import cellname

def getOrCreateSheet(sheetsDocument, sheetName):
    try:
        return sheetsDocument.getSheets().getByName(sheetName)
    except NoSuchElementException:
        # Must not currently be a burndown table sheet
        namedSheet = sheetsDocument.createInstance(
            'com.sun.star.sheet.Spreadsheet')
        sheetsDocument.getSheets().insertByName(sheetName, namedSheet)
        return namedSheet

def clearSheet(sheet):
    flags = CellFlags.VALUE | CellFlags.DATETIME | CellFlags.STRING \
        | CellFlags.HARDATTR | CellFlags.STYLES | CellFlags.OBJECTS \
        | CellFlags.EDITATTR | CellFlags.FORMATTED
    sheet.clearContents(flags)

class EmptyFormError(Exception):
    pass

class SheetRecord:
    def __init__(self, headers, record): # headers: {}, record: CellArray
        self.headers = headers
        self.record = record

    def __iter__(self):
        return iter(self.record)

    def getCount(self):
        return self.record.getCount()

    def getItem(self, index):
        return self.record.getItem(index)

    def __getitem__(self, arg):
        return self.record.getItem(self.headers[arg])

class SheetTableIterator:
    def __init__(self, headers, inner):
        self.headers = headers
        self.inner = inner

    def __next__(self):
        return SheetRecord(self.headers, next(self.inner))

class SheetTable:
    """Represents a group of non-empty cells with named columns"""
    def __init__(self, topLeft, tableWidth, xSheet):
        # Get the coordinates (in zero-indexed form)
        topLeftColumn, topLeftRow = cellname.getCoordinatesFromCellName(
            topLeft)

        # Get the number of rows in the table
        numberOfRows = 0
        dataProbeSpec = (
            f'{topLeft}:' + cellname.getCellNameFromCoordinates(
                topLeftColumn, cellname.ROW_MAX)
        )
        for row in CellMatrix(spec=dataProbeSpec, xSheet=xSheet):
            if not row.getItem(0).String:
                break
            numberOfRows += 1
        if not numberOfRows:
            raise EmptyFormError()

        # Grab headers and instantiate an inner container
        self._parseHeaders(topLeft, tableWidth, xSheet)
        rightColumn = topLeftColumn + tableWidth - 1
        dataSpec = (
            cellname.getCellNameFromCoordinates(topLeftColumn, topLeftRow + 1)
            + ':' + cellname.getCellNameFromCoordinates(
                rightColumn, numberOfRows + topLeftRow - 1)
        )
        self.container = CellMatrix(dataSpec, xSheet)

    def _parseHeaders(self, topLeft, tableWidth, xSheet):
        leftColumn, leftRow = cellname.getCoordinatesFromCellName(topLeft)
        rightColumn = leftColumn + tableWidth - 1
        headerSpec = (
            f'{topLeft}:'
            + cellname.getCellNameFromCoordinates(rightColumn, leftRow)
        )

        index = 0
        headerMap = {}
        row = CellMatrix(spec=headerSpec, xSheet=xSheet).getItem(0)
        for cell in row:
            headerMap[cell.String] = index
            index += 1
        self.headers = headerMap

    def __iter__(self):
        return SheetTableIterator(self.headers, iter(self.container))

    def getCount(self):
        return self.container.getCount()

    def getItem(self, index):
        return self.container.getItem(index)

    def getHeaders(self):
        return list(self.headers.keys())

###############################################################################
