###############################################################################
# NAME:             burndown.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Burndown chart/table logic
#
# CREATED:          12/03/2021
#
# LAST EDITED:      02/03/2022
###

from datetime import datetime
from typing import List

from .account import Account, AccountHistorySummaryForm
from .cellformat import NumberFormat
from .cellname import getCellNameFromCoordinates
from .cellrange import CellMatrix

class BurndownCalculator:
    def __init__(self, transactions, balances: AccountHistorySummaryForm,
                 burndownTableSheet, config):
        self.transactions = transactions
        self.balances = balances
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
    def getAccounts(balances: AccountHistorySummaryForm):
        """Assemble dictionary of accounts"""
        accounts = list(map(
            lambda a: Account(a.getName(), a.getStartingBalance()),
            balances.read()))
        accountNames = map(lambda a: a.getName(), accounts)
        return dict(zip(accountNames, accounts))

    @staticmethod
    def writeHeaders(outputIter, headers):
        """Write out Date, Description, & balance for each account for each
        transation in the ledger"""
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
        next(initialBalanceIter)
        next(initialBalanceIter)
        for account in accounts:
            cell = next(initialBalanceIter)
            cell.NumberFormat = NumberFormat.CURRENCY
            cell.Value = accounts[account].getBalance()
        for _, totaledAccounts in totals.items():
            totalBalance = 0
            for account in totaledAccounts:
                totalBalance += accounts[account].getBalance()
            cell = next(initialBalanceIter)
            cell.NumberFormat = NumberFormat.CURRENCY
            cell.Value = totalBalance
        return outputIter

    @staticmethod
    def writeBurndownTable(outputIter, transactions, accounts, totals,
                           startDate, endDate):
        """Write transactions to burndown table"""
        for transaction in transactions:
            if transaction.date < datetime.strptime(startDate, '%m/%d/%y'):
                continue
            if transaction.date > datetime.strptime(endDate, '%m/%d/%y'):
                break

            entry = iter(next(outputIter))
            next(entry).String = transaction.date.strftime('%m/%d/%y')
            next(entry).String = transaction.description
            next(entry).Value = transaction.amount
            mutatedAccountName = transaction.accountName
            next(entry).String = mutatedAccountName
            for name, account in accounts.items():
                cell = next(entry)
                if name == mutatedAccountName:
                    transaction.applyToAccount(account)
                cell.NumberFormat = NumberFormat.CURRENCY
                cell.Value = account.getBalance()

            for _, totaledAccounts in totals.items():
                totalBalance = 0
                for account in totaledAccounts:
                    totalBalance += accounts[account].getBalance()
                cell = next(entry)
                cell.NumberFormat = NumberFormat.CURRENCY
                cell.Value = totalBalance

    @staticmethod
    def writeFinalBalances(balances: AccountHistorySummaryForm,
                           accounts: List[Account]):
        summaries = list(map(lambda a: a.getHistorySummary(),
                             accounts.values()))
        balances.write(summaries)

    def run(self, startDate, endDate):
        accounts = BurndownCalculator.getAccounts(self.balances)
        headers = ['Date', 'Description', 'Amount', 'Account',
                   *accounts.keys(), *self.totals.keys()]
        numberOfEntries = len(self.transactions) + 1
        bottomCorner = getCellNameFromCoordinates(
            len(headers) - 1, numberOfEntries)
        burndownTable = CellMatrix(f'A1:{bottomCorner}',
                                   self.burndownTableSheet)

        # Write headers and initial balances
        outputIter = BurndownCalculator.writeInitialBalances(
            BurndownCalculator.writeHeaders(iter(burndownTable), headers),
            startDate, accounts, self.totals)

        BurndownCalculator.writeBurndownTable(
            outputIter, self.transactions, accounts, self.totals, startDate,
            endDate)
        # Write final balances
        BurndownCalculator.writeFinalBalances(self.balances, accounts)

###############################################################################
