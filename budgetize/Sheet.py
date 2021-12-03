###############################################################################
# NAME:             Sheet.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Abstractions for handling sheets
#
# CREATED:          12/02/2021
#
# LAST EDITED:      12/02/2021
###

from .CellRange import CellArray, CellMatrix, ROW_MAX, \
    getCoordinatesForCellSpec, getColumnNameFromIndex

class EmptyFormError(Exception):
    pass

# TODO: Custom iterator class that allows accessing via column header

class SheetTable:
    """Represents a group of non-empty cells with named columns"""
    def __init__(self, topLeft, rightmostColumn, xSheet):
        # Get the coordinates (in zero-indexed form)
        topLeftColumn, topLeftRow = getCoordinatesForCellSpec(topLeft)

        # Recondition them back to actual table coordinates
        topLeftColumnName = getColumnNameFromIndex(topLeftColumn)
        topLeftRow += 1

        # Get the number of rows in the table
        numberOfRows = 0
        dataProbeSpec = (
            f'{topLeftColumnName}{topLeftRow}'
            + f':{topLeftColumnName}{ROW_MAX}'
        )
        for cell in CellArray(spec=dataProbeSpec, xSheet=xSheet):
            if not cell.String:
                break
            numberOfRows += 1
        if not numberOfRows:
            raise EmptyFormError()

        # Grab headers and instantiate an inner container
        headerSpec = f'{topLeft}:{rightmostColumn}{topLeftRow}'
        self.headers = CellArray(spec=headerSpec, xSheet=xSheet)
        dataSpec = (
            f'{topLeftColumnName}{topLeftRow + 1}'
            + f':{rightmostColumn}{numberOfRows + topLeftRow - 1}'
        )
        self.container = CellMatrix(dataSpec, xSheet)

    def __iter__(self):
        return iter(self.container)

###############################################################################
