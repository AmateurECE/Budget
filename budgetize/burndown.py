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
                 burndownTableSheet):
        self.transactions = transactions
        self.balances = balances
        self.burndownTableSheet = burndownTableSheet

    @staticmethod
    def getAccounts(balances: AccountHistorySummaryForm):
        """Assemble dictionary of accounts"""
        accounts = list(map(
            lambda a: Account(a.getName(), a.getStartingBalance()),
            balances.read()))
        accountNames = map(lambda a: a.getName(), accounts)
        return dict(zip(accountNames, accounts))

    @staticmethod
    def getBurndownEntries(transactions, accounts, startDate, endDate):
        """Write transactions to burndown table"""
        entries = []
        for transaction in transactions:
            if transaction.date < datetime.strptime(startDate, '%m/%d/%y'):
                continue
            if transaction.date > datetime.strptime(endDate, '%m/%d/%y'):
                break

            affectedAccount = accounts[transaction.accountName]
            transaction.applyToAccount(affectedAccount)

            balances = list(map(lambda a: a.getBalance(), accounts.values()))
            entries.append(BurndownEntry(transaction, balances))
        return entries

    @staticmethod
    def writeFinalBalances(balances: AccountHistorySummaryForm,
                           accounts: List[Account]):
        summaries = list(map(lambda a: a.getHistorySummary(),
                             accounts.values()))
        balances.write(summaries)

    def run(self, startDate, endDate):
        accounts = BurndownCalculator.getAccounts(self.balances)
        bottomCorner = getCellNameFromCoordinates(
            4 + len(accounts) + - 1,
            len(self.transactions) + 1)

        table = CellMatrix(f'A1:{bottomCorner}', self.burndownTableSheet)
        entries = BurndownCalculator.getBurndownEntries(
            self.transactions, accounts, startDate, endDate)
        BurndownForm(table, accounts).write(startDate, entries)
        BurndownCalculator.writeFinalBalances(self.balances, accounts)

###############################################################################
# BurndownForm
###

class BurndownEntry:
    def __init__(self, transaction, balances):
        self.date = transaction.date
        self.description = transaction.description
        self.amount = transaction.amount
        self.accountName = transaction.accountName
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
        next(iterator).String = datetime.strftime(entry.getDate(), '%m/%d/%y')
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
    def writeInitialBalances(iterator, startDate, initialBalances):
        next(iterator).String = startDate
        next(iterator).String = 'Starting Balance'
        next(iterator)
        next(iterator)
        for balance in initialBalances:
            balanceField = next(iterator)
            balanceField.Value = balance
            balanceField.NumberFormat = NumberFormat.CURRENCY

    def write(self, startDate, entries: List[BurndownRecord]):
        headers = ['Date', 'Description', 'Amount', 'Account',
                   *self.startingBalances.keys()]
        iterator = iter(self.cellrange)
        BurndownForm.writeHeaders(iter(next(iterator)), headers)
        BurndownForm.writeInitialBalances(
            iter(next(iterator)), startDate, self.startingBalances.values())

        for entry in entries:
            BurndownRecord(next(iterator)).write(entry)

###############################################################################
