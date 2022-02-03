###############################################################################
# NAME:             account.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Account classes
#
# CREATED:          12/03/2021
#
# LAST EDITED:      02/03/2022
###

from typing import List

from .cellrange import CellRow, CellMatrix
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
    def __init__(self, name, startingBalance, endingBalance):
        self.name = name
        self.startingBalance = startingBalance
        self.endingBalance = endingBalance

    def getName(self):
        return self.name

    def getStartingBalance(self):
        return self.startingBalance

    def getEndingBalance(self):
        return self.endingBalance

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

        endingBalanceCell = next(iterator)
        endingBalanceCell.NumberFormat = NumberFormat.CURRENCY
        endingBalanceCell.Value = summary.getEndingBalance()

class AccountHistorySummaryForm:
    """Deals with persistence of a list of AccountHistorySummary instances"""
    def __init__(self, cellrange: SheetTable):
        self.cellrange = cellrange
        headers = cellrange.getHeaders()
        self.startDate = headers[1]
        self.endDate = headers[2]

    def getStartDate(self):
        return self.startDate

    def getEndDate(self):
        return self.endDate

    def read(self) -> List[AccountHistorySummary]:
        result = []
        for row in self.cellrange:
            result.append(AccountHistorySummaryRecord(row).read())
        return result

    def write(self, summaries: List[AccountHistorySummary]):
        iterator = iter(self.cellrange)
        for summary in summaries:
            AccountHistorySummaryRecord(next(iterator)).write(summary)

###############################################################################
