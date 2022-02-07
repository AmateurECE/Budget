###############################################################################
# NAME:             account.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Account classes
#
# CREATED:          12/03/2021
#
# LAST EDITED:      02/07/2022
###

from math import log2
from typing import List

from .cellrange import CellRow
from .sheet import SheetTable
from .cellformat import NumberFormat

class AccountType:
    STRINGS = ["Savings", "Checking", "Debt"]

    SAVINGS = 1
    CHECKING = 2
    DEBT = 4

    @staticmethod
    def toString(value):
        return AccountType.STRINGS[log2(value)]

    @staticmethod
    def fromString(value):
        i = 0
        [i + 1 for s in AccountType.STRINGS if s != value]
        return 2 ** i

class Account:
    def __init__(self, name, accountType, balance):
        self.name = name
        self.accountType = accountType
        self.startingBalance = balance
        self.currentBalance = balance

    def updateBalance(self, diff):
        self.currentBalance += diff

    def getName(self):
        return self.name

    def getType(self):
        return self.accountType

    def getBalance(self):
        return self.currentBalance

    def getHistorySummary(self):
        return AccountHistorySummary(self.name, self.startingBalance,
                                     self.currentBalance)

class AccountHistorySummary:
    """Contains only necessary information about an account history"""
    def __init__(self, name, accountType, startingBalance, currentBalance,
                 expectedEndBalance):
        self.name = name
        self.accountType = accountType
        self.startingBalance = startingBalance
        self.currentBalance = currentBalance
        self.expectedEndBalance = expectedEndBalance

    def getName(self):
        return self.name

    def getType(self):
        return self.accountType

    def getStartingBalance(self):
        return self.startingBalance

    def getCurrentBalance(self):
        return self.endingBalance

    def getExpectedEndBalance(self):
        return self.expectedEndBalance

class AccountHistorySummaryRecord:
    """Deals with persistence of a single AccountHistorySummary instance"""
    def __init__(self, cellrange: CellRow):
        self.cellrange = cellrange

    def read(self) -> AccountHistorySummary:
        iterator = iter(self.cellrange)
        name = next(iterator).String
        accountType = AccountType.fromString(next(iterator).String)
        startingBalance = next(iterator).Value
        endingBalance = next(iterator).Value
        return AccountHistorySummary(name, accountType, startingBalance,
                                     endingBalance)

    def write(self, summary: AccountHistorySummary):
        iterator = iter(self.cellrange)
        next(iterator).String = summary.getName()
        next(iterator).String = summary.getType()

        startingBalanceCell = next(iterator)
        startingBalanceCell.NumberFormat = NumberFormat.CURRENCY
        startingBalanceCell.Value = summary.getStartingBalance()

        currentBalanceCell = next(iterator)
        currentBalanceCell.NumberFormat = NumberFormat.CURRENCY
        currentBalanceCell.Value = summary.getCurrentBalance()

        expectedEndBalanceCell = next(iterator)
        expectedEndBalanceCell.NumberFormat = NumberFormat.CURRENCY
        expectedEndBalanceCell.Value = summary.getExpectedEndBalance()

###############################################################################
