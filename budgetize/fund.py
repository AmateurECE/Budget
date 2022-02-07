###############################################################################
# NAME:             fund.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Sinking fund logic
#
# CREATED:          02/07/2022
#
# LAST EDITED:      02/07/2022
###

from .cellformat import NumberFormat
from .cellrange import CellRow

class SinkingFund:
    def __init__(self, description, accountName, startingBalance,
                 currentBalance, expectedEndBalance):
        self.description = description
        self.accountName = accountName
        self.startingBalance = startingBalance
        self.currentBalance = currentBalance
        self.expectedEndBalance = expectedEndBalance

    def getDescription(self):
        return self.description

    def getAccountName(self):
        return self.accountName

    def getStartingBalance(self):
        return self.startingBalance

    def getCurrentBalance(self):
        return self.currentBalance

    def getExpectedEndBalance(self):
        return self.expectedEndBalance

class SinkingFundRecord:
    def __init__(self, cellrange: CellRow):
        self.cellrange = cellrange

    def write(self, fund: SinkingFund):
        iterator = iter(self.cellrange)
        next(iterator).String = fund.getDescription()
        next(iterator).String = fund.getAccountName()

        startingBalanceCell = next(iterator)
        startingBalanceCell.NumberFormat = NumberFormat.CURRENCY
        startingBalanceCell.Value = fund.getStartingBalance()

        currentBalanceCell = next(iterator)
        currentBalanceCell.NumberFormat = NumberFormat.CURRENCY
        currentBalanceCell.Value = fund.getCurrentBalance()

        expectedEndBalanceCell = next(iterator)
        expectedEndBalanceCell.NumberFormat = NumberFormat.CURRENCY
        expectedEndBalanceCell.Value = fund.getExpectedEndBalance()

    def read(self) -> SinkingFund:
        recordIterator = iter(self.cellrange)
        description = next(recordIterator).String
        accountName = next(recordIterator).String
        startingBalance = next(recordIterator).Value
        endingBalance = next(recordIterator).Value
        expectedEndBalance = next(recordIterator).Value
        return SinkingFund(description, accountName, startingBalance,
                           endingBalance, expectedEndBalance)

###############################################################################
