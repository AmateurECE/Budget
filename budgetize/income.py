###############################################################################
# NAME:             income.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Types for handling incomes.
#
# CREATED:          02/06/2022
#
# LAST EDITED:      02/07/2022
###

from .cellformat import NumberFormat
from .cellrange import CellRow

class Income:
    def __init__(self, description: str, accountName: str, amount: float):
        self.description = description
        self.accountName = accountName
        self.amount = amount
        self.received = 0.0

    def getDescription(self):
        return self.description

    def getAccountName(self):
        return self.accountName

    def getAmount(self):
        return self.amount

    def receive(self, amount):
        self.received += amount

    def getReceived(self):
        return self.received

class IncomeRecord:
    def __init__(self, cellrange: CellRow):
        self.cellrange = cellrange

    def read(self) -> Income:
        recordIterator = iter(self.cellrange)
        description = next(recordIterator).String
        accountName = next(recordIterator).String
        amount = next(recordIterator).Value
        return Income(description, accountName, amount)

    def write(self, income: Income):
        recordIterator = iter(self.cellrange)
        next(recordIterator).String = income.getDescription()
        next(recordIterator).String = income.getAccountName()
        expectedCell = next(recordIterator)
        expectedCell.NumberFormat = NumberFormat.CURRENCY
        expectedCell.Value = income.getAmount()

        receivedCell = next(recordIterator)
        receivedCell.NumberFormat = NumberFormat.CURRENCY
        receivedCell.Value = income.getReceived()

###############################################################################
