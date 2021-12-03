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

from . import CellRange

class Budgetizer:
    def __init__(self, xSheetDoc):
        self.sheetDoc = xSheetDoc

    @staticmethod
    def getNonRecurringForm(xSheet):
        numberOfRows = 0
        for cell in CellRange.getCellRangeForListSpec(
                f'A1:A{CellRange.ROW_MAX}', xSheet):
            if not cell.String:
                break
            numberOfRows += 1
        return CellRange.getCellRangeForMatrixSpec(
            f'A2:D{numberOfRows}', xSheet)

    @staticmethod
    def getBeginningBalances(xSheet):
        numberOfRows = 0
        for cell in CellRange.getCellRangeForListSpec(
                f'A3:A{CellRange.ROW_MAX}', xSheet):
            if not cell.String:
                break
            numberOfRows += 1
        if not numberOfRows:
            raise RuntimeError('No accounts have been configured!')
        return CellRange.getCellRangeForMatrixSpec(
            f'A3:B{2 + numberOfRows}', xSheet)

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
            accountName = account.xIndexAccess.getCellByPosition(0, 0)
            accountBalance = account.xIndexAccess.getCellByPosition(1, 0)
            print(f'{accountName.String}: {accountBalance.String}')

###############################################################################
