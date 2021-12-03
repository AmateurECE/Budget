###############################################################################
# NAME:             Budgetizer.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Budgetizer entrypoint
#
# CREATED:          12/01/2021
#
# LAST EDITED:      12/02/2021
###

from .CellRange import CellArray, CellMatrix, ROW_MAX
from .Sheet import SheetTable

class Budgetizer:
    def __init__(self, xSheetDoc):
        self.sheetDoc = xSheetDoc

    def budgetize(self):
        nonRecurringForm = SheetTable(
            "A1", 4, self.sheetDoc.getSheets().getByName("Non Recurring"))
        beginningBalances = SheetTable(
            "A2", 2, self.sheetDoc.getSheets().getByName("Balances"))
        for record in nonRecurringForm:
            for cell in record:
                print(cell.String + ',', end='')
            print()
        for account in beginningBalances:
            accountName = account.getItem(0).String
            accountBalance = account.getItem(1).String
            print(f'{accountName}: {accountBalance}')

###############################################################################
