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
    getCoordinatesFromCellName, getColumnNameFromIndex, \
    getCellNameFromCoordinates

class EmptyFormError(Exception):
    pass

# TODO: Custom iterator class that allows accessing via column header
# TODO: Move Cell/coordinates logic into new module

class SheetTable:
    """Represents a group of non-empty cells with named columns"""
    def __init__(self, topLeft, tableWidth, xSheet):
        # Get the coordinates (in zero-indexed form)
        topLeftColumn, topLeftRow = getCoordinatesFromCellName(topLeft)

        # Get the number of rows in the table
        numberOfRows = 0
        dataProbeSpec = (
            f'{topLeft}:' + getCellNameFromCoordinates(topLeftColumn, ROW_MAX)
        )
        for cell in CellArray(spec=dataProbeSpec, xSheet=xSheet):
            if not cell.String:
                break
            numberOfRows += 1
        if not numberOfRows:
            raise EmptyFormError()

        # Grab headers and instantiate an inner container
        rightColumn = topLeftColumn + tableWidth - 1
        headerSpec = (
            f'{topLeft}:'
            + getCellNameFromCoordinates(rightColumn, topLeftRow)
        )
        self.headers = CellArray(spec=headerSpec, xSheet=xSheet)
        dataSpec = (
            getCellNameFromCoordinates(topLeftColumn, topLeftRow + 1)
            + ':' + getCellNameFromCoordinates(
                rightColumn, numberOfRows + topLeftRow - 1)
        )
        self.container = CellMatrix(dataSpec, xSheet)

    def __iter__(self):
        return iter(self.container)

###############################################################################
