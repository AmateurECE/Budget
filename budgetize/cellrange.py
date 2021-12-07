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
# LAST EDITED:      12/06/2021
###

from . import cellname

RowAccessor = lambda i, dim, access: access.getCellByPosition(i, dim)

class CellRowIterator:
    def __init__(self, row, columns, xIndexAccess):
        self.row = row
        self.columns = columns
        self.index = 0
        self.xIndexAccess = xIndexAccess

    def __next__(self):
        if self.index >= self.columns:
            raise StopIteration()
        index = self.index
        self.index += 1
        return RowAccessor(index, self.row, self.xIndexAccess)

class CellRow:
    def __init__(self, **kwargs):
        """Two forms, really:
        1. spec, accessor
        2. row, columns, accessor
        """
        if 'spec' in kwargs.keys():
            self.initString(kwargs)
        else:
            self.initIndex(kwargs)

    def initString(self, kwargs):
        spec = kwargs.get('spec', None)
        firstSpec, secondSpec = tuple(spec.split(':'))
        firstColumn, firstRow = cellname.getCoordinatesFromCellName(firstSpec)
        secondColumn, secondRow = cellname.getCoordinatesFromCellName(
            secondSpec)
        self.xIndexAccess = kwargs.get(
            'accessor', None).getCellRangeByName(spec)
        if firstRow != secondRow:
            raise RuntimeError(f'Poorly formed spec ({spec}) for CellRow!')
        self.columns = secondColumn - firstColumn + 1
        self.row = firstRow

    def initIndex(self, kwargs):
        self.row = kwargs.get('row')
        self.columns = kwargs.get('columns', None)
        self.xIndexAccess = kwargs.get('accessor', None)

    def __iter__(self):
        return CellRowIterator(self.row, self.columns, self.xIndexAccess)

    def getCount(self):
        return self.columns

    def getRow(self):
        return self.row

    def getItem(self, index):
        return RowAccessor(index, self.row, self.xIndexAccess)

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
        return CellRow(row=index, columns=self.columns,
                       accessor=self.xIndexAccess)

class CellMatrix:
    def __init__(self, spec, xSheet):
        firstSpec, secondSpec = tuple(spec.split(':'))
        firstColumn, firstRow = cellname.getCoordinatesFromCellName(firstSpec)
        secondColumn, secondRow = cellname.getCoordinatesFromCellName(
            secondSpec)
        self.xIndexAccess = xSheet.getCellRangeByName(spec)
        if firstRow == secondRow:
            self.rows = 1
            self.columns = secondColumn - firstColumn + 1
        elif firstColumn == secondColumn:
            self.rows = secondRow - firstRow + 1
            self.columns = 1
        else:
            self.rows = secondRow - firstRow + 1
            self.columns = secondColumn - firstColumn + 1

    def __iter__(self):
        return CellMatrixIterator(self.rows, self.columns, self.xIndexAccess)

    def getCount(self):
        return self.rows

    def getItem(self, index):
        return CellRow(row=index, columns=self.columns,
                       accessor=self.xIndexAccess)

###############################################################################
