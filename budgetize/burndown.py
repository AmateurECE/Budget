###############################################################################
# NAME:             burndown.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Burndown chart/table logic
#
# CREATED:          12/03/2021
#
# LAST EDITED:      12/03/2021
###

from .account import Account
from .cellname import getCellNameFromCoordinates
from .cellrange import CellMatrix

class BurndownCalculator:
    def __init__(
            self, nonRecurringForm, beginningBalances, burndownTableSheet):
        self.nonRecurringForm = nonRecurringForm
        self.beginningBalances = beginningBalances
        self.burndownTableSheet = burndownTableSheet

    @staticmethod
    def getAccounts(beginningBalances, startDate):
        """Assemble dictionary of accounts"""
        accounts = list(map(
            lambda a: Account(a['Account'].String, a[startDate].Value),
            beginningBalances))
        accountNames = map(lambda a: a.getName(), accounts)
        return dict(zip(accountNames, accounts))

    @staticmethod
    def writeHeaders(outputIter, headers):
        """Write out Date, Description, & balance for each account for each
        transation in the nonRecurringForm"""
        headerRowIter = iter(next(outputIter))
        for header in headers:
            next(headerRowIter).String = header
        return outputIter

    @staticmethod
    def writeInitialBalances(outputIter, startDate, accounts):
        initialBalance = next(outputIter)
        initialBalanceIter = iter(initialBalance)
        next(initialBalanceIter).String = startDate
        next(initialBalanceIter).String = 'Starting Balance'
        for account in accounts:
            next(initialBalanceIter).Value = accounts[account].getBalance()
        return outputIter

    @staticmethod
    def writeBurndownTable(nonRecurringForm, outputIter, accounts):
        """Write transactions to burndown table"""
        for transaction in nonRecurringForm:
            entry = iter(next(outputIter))
            account = accounts[transaction['Account'].String]
            account.updateBalance(transaction['Amount'].Value)
            next(entry).String = transaction['Date'].String
            next(entry).String = transaction['Description'].String
            for account in accounts:
                next(entry).Value = accounts[account].getBalance()

    @staticmethod
    def writeFinalBalances(beginningBalances, accounts):
        endDate = beginningBalances.getHeaders()[2]
        for record in beginningBalances:
            record[endDate].Value = accounts[
                record['Account'].String].getBalance()

    def run(self):
        startDate = self.beginningBalances.getHeaders()[1]
        accounts = BurndownCalculator.getAccounts(self.beginningBalances,
                                                  startDate)
        headers = ['Date', 'Description', *accounts.keys()]
        numberOfEntries = self.nonRecurringForm.getCount() + 1
        bottomCorner = getCellNameFromCoordinates(
            len(headers) - 1, numberOfEntries)
        burndownTable = CellMatrix(f'A1:{bottomCorner}',
                                   self.burndownTableSheet)

        # Write headers and initial balances
        outputIter = BurndownCalculator.writeInitialBalances(
            BurndownCalculator.writeHeaders(iter(burndownTable), headers),
            startDate, accounts)
        BurndownCalculator.writeBurndownTable(self.nonRecurringForm,
                                              outputIter, accounts)
        # Write final balances
        BurndownCalculator.writeFinalBalances(self.beginningBalances, accounts)

###############################################################################
