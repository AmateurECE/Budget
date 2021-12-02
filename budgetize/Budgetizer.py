###############################################################################
# NAME:             Budgetizer.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Budgetizer entrypoint
#
# CREATED:          12/01/2021
#
# LAST EDITED:      12/01/2021
###

from .CellRange import getCellRangeForMatrixSpec

class Budgetizer:
    def __init__(self, xSheetDocController):
        self.sheetDocController = xSheetDocController

    def budgetize(self):
        xSheet = self.sheetDocController.ActiveSheet
        matrix = getCellRangeForMatrixSpec("A1:B2", xSheet)
        for record in matrix:
            for cell in record:
                cell.String = 'Budgetize!'

###############################################################################
