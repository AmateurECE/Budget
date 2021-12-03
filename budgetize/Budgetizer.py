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

class Budgetizer:
    def __init__(self, xSheetDoc):
        self.sheetDoc = xSheetDoc

    @staticmethod
    def getNonRecurringForm(xSheet):
        numberOfRows = 0
        for cell in CellArray(spec=f'A1:A{ROW_MAX}', xSheet=xSheet):
            if not cell.String:
                break
            numberOfRows += 1
        return CellMatrix(f'A2:D{numberOfRows}', xSheet)

    @staticmethod
    def getBeginningBalances(xSheet):
        numberOfRows = 0
        for cell in CellArray(spec=f'A3:A{ROW_MAX}', xSheet=xSheet):
            if not cell.String:
                break
            numberOfRows += 1
        if not numberOfRows:
            raise RuntimeError('No accounts have been configured!')
        return CellMatrix(f'A3:B{2 + numberOfRows}', xSheet)

    def budgetize(self):
        nonRecurringForm = Budgetizer.getNonRecurringForm(
            self.sheetDoc.getSheets().getByName("Non Recurring"))
        beginningBalances = Budgetizer.getBeginningBalances(
            self.sheetDoc.getSheets().getByName("Balances"))
        for record in nonRecurringForm:
            for cell in record:
                print(cell.String + ',', end='')
            print()
        for account in beginningBalances:
            accountName = account.getItem(0).String
            accountBalance = account.getItem(1).String
            print(f'{accountName}: {accountBalance}')

###############################################################################
