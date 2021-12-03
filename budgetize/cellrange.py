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

from . import cellname

RowIteratorFn = lambda index, dim, access: access.getCellByPosition(index, dim)
ColumnIteratorFn = lambda index, dim, access: access.getCellByPosition(
    dim, index)

class CellListIterator:
    def __init__(self, dim, maxIndex, xIndexAccess, iteratorFn):
        self.dim = dim
        self.maxIndex = maxIndex
        self.index = 0
        self.xIndexAccess = xIndexAccess
        self.iteratorFn = iteratorFn

    def __next__(self):
        if self.index >= self.maxIndex:
            raise StopIteration()
        index = self.index
        self.index += 1
        return self.iteratorFn(index, self.dim, self.xIndexAccess)

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
        return CellArray(
            dimension=index, rows=self.columns, accessor=self.xIndexAccess,
            iteratorFn=RowIteratorFn
        )

class CellArray:
    def __init__(self, **kwargs):
        """Two forms, really:
        1. spec, xSheet, dimension(?)
        2. columns, accessor, iteratorFn, dimension
        """
        self.dimension = kwargs.get('dimension', 0)
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
        self.xIndexAccess = kwargs.get('xSheet', None).getCellRangeByName(spec)
        if firstRow == secondRow:
            self.rows = secondColumn - firstColumn + 1
            self.iteratorFn = RowIteratorFn
        elif firstColumn == secondColumn:
            self.rows = secondRow - firstRow + 1
            self.iteratorFn = ColumnIteratorFn
        else:
            raise RuntimeError(f'Poorly formed spec ({spec}) for list!')

    def initIndex(self, kwargs):
        self.rows = kwargs.get('rows', None)
        self.xIndexAccess = kwargs.get('accessor', None)
        self.iteratorFn = kwargs.get('iteratorFn', None)

    def __iter__(self):
        return CellListIterator(
            self.dimension, self.rows, self.xIndexAccess, self.iteratorFn)

    def getCount(self):
        return self.rows

    def getItem(self, index):
        return self.iteratorFn(index, self.dimension, self.xIndexAccess)

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
        raise NotImplementedError()

###############################################################################
