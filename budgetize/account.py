###############################################################################
# NAME:             account.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Account classes
#
# CREATED:          12/03/2021
#
# LAST EDITED:      02/06/2022
###

from typing import List

from .cellrange import CellRow
from .sheet import SheetTable
from .cellformat import NumberFormat

class Account:
    def __init__(self, name, balance):
        self.name = name
        self.startingBalance = balance
        self.currentBalance = balance

    def updateBalance(self, diff):
        self.currentBalance += diff

    def getName(self):
        return self.name

    def getBalance(self):
        return self.currentBalance

    def getHistorySummary(self):
        return AccountHistorySummary(self.name, self.startingBalance,
                                     self.currentBalance)

class AccountHistorySummary:
    """Contains only necessary information about an account history"""
    def __init__(self, name, startingBalance, currentBalance,
                 expectedEndBalance):
        self.name = name
        self.startingBalance = startingBalance
        self.currentBalance = currentBalance
        self.expectedEndBalance = expectedEndBalance

    def getName(self):
        return self.name

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
        startingBalance = next(iterator).Value
        endingBalance = next(iterator).Value
        return AccountHistorySummary(name, startingBalance, endingBalance)

    def write(self, summary: AccountHistorySummary):
        iterator = iter(self.cellrange)
        next(iterator).String = summary.getName()

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
