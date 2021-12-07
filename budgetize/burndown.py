###############################################################################
# NAME:             burndown.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Burndown chart/table logic
#
# CREATED:          12/03/2021
#
# LAST EDITED:      12/06/2021
###

from .account import Account
from .cellformat import NumberFormat
from .cellname import getCellNameFromCoordinates
from .cellrange import CellMatrix

class BurndownCalculator:
    def __init__(self, nonRecurringForm, beginningBalances,
                 burndownTableSheet, config):
        self.nonRecurringForm = nonRecurringForm
        self.beginningBalances = beginningBalances
        self.burndownTableSheet = burndownTableSheet
        self.totals = {}
        for row in config:
            self.parseConfigRow(row)

    MAX_CONFIG_COLUMNS = 2
    def parseConfigRow(self, row):
        """Parse a single row of configuration"""
        firstColumn = row.getItem(0).String
        if firstColumn.startswith('total:'):
            secondColumn = row.getItem(1).String
            self.totals[firstColumn.replace('total:', '')] = \
                secondColumn.split(',')

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
    def writeInitialBalances(outputIter, startDate, accounts, totals):
        initialBalance = next(outputIter)
        initialBalanceIter = iter(initialBalance)
        next(initialBalanceIter).String = startDate
        next(initialBalanceIter).String = 'Starting Balance'
        for account in accounts:
            cell = next(initialBalanceIter)
            cell.NumberFormat = NumberFormat.CURRENCY
            cell.Value = accounts[account].getBalance()
        for totalName, totaledAccounts in totals.items():
            totalBalance = 0
            for account in totaledAccounts:
                totalBalance += accounts[account].getBalance()
            cell = next(initialBalanceIter)
            cell.NumberFormat = NumberFormat.CURRENCY
            cell.Value = totalBalance
        return outputIter

    @staticmethod
    def writeBurndownTable(nonRecurringForm, outputIter, accounts, totals):
        """Write transactions to burndown table"""
        for transaction in nonRecurringForm:
            entry = iter(next(outputIter))
            account = accounts[transaction['Account'].String]
            account.updateBalance(transaction['Amount'].Value)
            next(entry).String = transaction['Date'].String
            next(entry).String = transaction['Description'].String
            for account in accounts:
                cell = next(entry)
                cell.NumberFormat = NumberFormat.CURRENCY
                cell.Value = accounts[account].getBalance()
            for totalName, totaledAccounts in totals.items():
                totalBalance = 0
                for account in totaledAccounts:
                    totalBalance += accounts[account].getBalance()
                cell = next(entry)
                cell.NumberFormat = NumberFormat.CURRENCY
                cell.Value = totalBalance

    @staticmethod
    def writeFinalBalances(beginningBalances, accounts):
        endDate = beginningBalances.getHeaders()[2]
        for record in beginningBalances:
            record[endDate].NumberFormat = NumberFormat.CURRENCY
            record[endDate].Value = accounts[
                record['Account'].String].getBalance()

    def run(self):
        startDate = self.beginningBalances.getHeaders()[1]
        accounts = BurndownCalculator.getAccounts(self.beginningBalances,
                                                  startDate)
        headers = ['Date', 'Description', *accounts.keys(),
                   *self.totals.keys()]
        numberOfEntries = self.nonRecurringForm.getCount() + 1
        bottomCorner = getCellNameFromCoordinates(
            len(headers) - 1, numberOfEntries)
        burndownTable = CellMatrix(f'A1:{bottomCorner}',
                                   self.burndownTableSheet)

        # Write headers and initial balances
        outputIter = BurndownCalculator.writeInitialBalances(
            BurndownCalculator.writeHeaders(iter(burndownTable), headers),
            startDate, accounts, self.totals)
        BurndownCalculator.writeBurndownTable(
            self.nonRecurringForm, outputIter, accounts, self.totals)
        # Write final balances
        BurndownCalculator.writeFinalBalances(self.beginningBalances, accounts)

###############################################################################
