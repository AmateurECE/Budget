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
from .cellrange import CellMatrix, CellRow

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
# BurndownForm
###

class BurndownEntry:
    def __init__(self, date, description, account, balances):
        self.date = date
        self.description = description
        self.amount = account.getBalance()
        self.accountName = account.getName()
        self.balances = balances

    def getDate(self):
        return self.date

    def getDescription(self):
        return self.description

    def getAmount(self):
        return self.amount

    def getAccountName(self):
        return self.accountName

    def getBalances(self):
        return self.balances

class BurndownRecord:
    def __init__(self, cellrange: CellRow):
        self.cellrange = cellrange

    def write(self, entry: BurndownEntry):
        iterator = iter(self.cellrange)
        next(iterator).String = datetime.strptime(entry.getDate(), '%m/%d/%y')
        next(iterator).String = entry.getDescription()
        amountField = next(iterator)
        amountField.Value = entry.getAmount()
        amountField.NumberFormat = NumberFormat.CURRENCY
        next(iterator).String = entry.getAccountName()

        for balance in entry.getBalances():
            balanceField = next(iterator)
            balanceField.Value = balance
            balanceField.NumberFormat = NumberFormat.CURRENCY

class BurndownForm:
    def __init__(self, cellrange: CellMatrix, accounts: dict[str, Account]):
        self.cellrange = cellrange
        self.startingBalances = {}
        for name in accounts:
            self.startingBalances[name] = accounts[name].getBalance()

    @staticmethod
    def writeHeaders(iterator, headers):
        for header in headers:
            next(iterator).String = header

    @staticmethod
    def writeInitialBalances(iterator, initialBalances):
        next(iterator).String = 'Starting Balance'
        next(iterator)
        next(iterator)
        for balance in initialBalances:
            balanceField = next(iterator)
            balanceField.Value = balance
            balanceField.NumberFormat = NumberFormat.CURRENCY

    def write(self, entries: List[BurndownRecord]):
        headers = ['Date', 'Description', 'Amount', 'Account',
                   *self.startingBalances.keys()]
        iterator = iter(self.cellrange)
        BurndownForm.writeHeaders(iter(next(iterator)), headers)
        BurndownForm.writeInitialBalances(
            iter(next(iterator)), self.startingBalances.values())

        for entry in entries:
            BurndownRecord(next(iterator)).write(entry)

###############################################################################
