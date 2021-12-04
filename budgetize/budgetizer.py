###############################################################################
# NAME:             Budgetizer.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Budgetizer entrypoint
#
# CREATED:          12/01/2021
#
# LAST EDITED:      12/03/2021
###

from .account import Account
from .cellname import getCellNameFromCoordinates
from .cellrange import CellMatrix
from .sheet import SheetTable

class Budgetizer:
    def __init__(self, xSheetDoc):
        self.sheetDoc = xSheetDoc

    def runBurndown(self, nonRecurringForm, beginningBalances):
        burndownTableSheetName = "Burndown Table"
        try:
            burndownTableSheet = self.sheetDoc.getSheets().getByName(
                burndownTableSheetName)
        except Exception:
            # Must not currently be a burndown table sheet
            burndownTableSheet = self.sheetDoc.createInstance(
                'com.sun.star.sheet.Spreadsheet')
            self.sheetDoc.getSheets().insertByName(
                burndownTableSheetName, burndownTableSheet)

        # Put togther dictionary of accounts
        startDate = beginningBalances.getHeaders()[1]
        accounts = list(map(
            lambda a: Account(a['Account'].String, a[startDate].Value),
            beginningBalances))
        accountNames = map(lambda a: a.getName(), accounts)
        accounts = dict(zip(accountNames, accounts))

        # Write out Date, Description, & balance for each account for each
        # transation in the nonRecurringForm
        numberOfEntries = nonRecurringForm.getCount() + 1
        headers = ['Date', 'Description', *accounts.keys()]
        bottomCorner = getCellNameFromCoordinates(
            len(headers) - 1, numberOfEntries)
        burndownTable = CellMatrix(f'A1:{bottomCorner}', burndownTableSheet)
        outputIter = iter(burndownTable)
        headerRowIter = iter(next(outputIter))
        for header in headers:
            next(headerRowIter).String = header
        endDate = beginningBalances.getHeaders()[2]
        initialBalance = next(outputIter)
        initialBalanceIter = iter(initialBalance)
        next(initialBalanceIter).String = startDate
        next(initialBalanceIter).String = 'Starting Balance'
        for account in accounts:
            next(initialBalanceIter).Value = accounts[account].getBalance()
        previousEntry = initialBalance
        for transaction in nonRecurringForm:
            entry = iter(next(outputIter))
            accountIndex = list(accounts.keys()).index(
                transaction['Account'].String)
            account = accounts[transaction['Account'].String]
            account.updateBalance(transaction['Amount'].Value)
            next(entry).String = transaction['Date'].String
            next(entry).String = transaction['Description'].String
            for account in accounts:
                next(entry).Value = accounts[account].getBalance()

    def budgetize(self):
        nonRecurringForm = SheetTable(
            "A1", 4, self.sheetDoc.getSheets().getByName("Non Recurring"))
        beginningBalances = SheetTable(
            "A1", 3, self.sheetDoc.getSheets().getByName("Balances"))
        self.runBurndown(nonRecurringForm, beginningBalances)

###############################################################################
