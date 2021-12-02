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

class Budgetizer:
    def __init__(self, xSheetDocController):
        self.sheetDocController = xSheetDocController

    def budgetize(self):
        xSheet = self.sheetDocController.ActiveSheet
        length = 2
        xCellRange = xSheet.getCellRangeByName(f'A1:A{length}')
        for i in range(0, length):
            xCellRange.getCellByPosition(0, i).String = 'Budgetize!'

###############################################################################
