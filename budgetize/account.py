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

from typing import List

from .cellrange import CellRow
from .sheet import SheetTable
from .cellformat import NumberFormat

class AccountHistorySummary:
    """Contains only necessary information about an account history"""
    def __init__(self, name, startingBalance):
        self.name = name
        self.startingBalance = startingBalance
        self.currentBalance = startingBalance
        self.expectedEndBalance = startingBalance

    def getAccountName(self):
        return self.name

    def getStartingBalance(self):
        return self.startingBalance

    def getExpectedEndBalance(self):
        return self.expectedEndBalance

    def getCurrentBalance(self):
        return self.currentBalance

    def updateBalance(self, diff):
        self.currentBalance += diff

    def updateExpectedBalance(self, diff):
        self.expectedEndBalance += diff

class AccountHistorySummaryRecord:
    """Deals with persistence of a single AccountHistorySummary instance"""
    def __init__(self, cellrange: CellRow):
        self.cellrange = cellrange

    def read(self) -> AccountHistorySummary:
        iterator = iter(self.cellrange)
        name = next(iterator).String
        startingBalance = next(iterator).Value
        return AccountHistorySummary(name, startingBalance)

    def write(self, summary: AccountHistorySummary):
        iterator = iter(self.cellrange)
        next(iterator).String = summary.getAccountName()

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
