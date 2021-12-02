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

from .CellRange import getCellRangeForMatrixSpec, ROW_MAX

class Budgetizer:
    def __init__(self, xSheetDoc):
        self.sheetDoc = xSheetDoc

    @staticmethod
    def getNonRecurringForm(xSheet):
        numberOfRows = 0
        for cell in getCellRangeForMatrixSpec(f'A1:A{ROW_MAX}', xSheet):
            if not cell.String:
                break
            numberOfRows += 1
        return getCellRangeForMatrixSpec(f'A2:D{numberOfRows}', xSheet)

    def budgetize(self):
        nonRecurringForm = Budgetizer.getNonRecurringForm(
            self.sheetDoc.getSheets().getByName("Non Recurring"))
        for record in nonRecurringForm:
            for cell in record:
                print(cell.String + ',', end='')
            print()

###############################################################################
