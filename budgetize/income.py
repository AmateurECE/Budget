###############################################################################
# NAME:             income.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Types for handling incomes.
#
# CREATED:          02/06/2022
#
# LAST EDITED:      02/06/2022
###

from .cellformat import NumberFormat
from .cellrange import CellRow

class Income:
    def __init__(self, description: str, amount: float):
        self.description = description
        self.amount = amount

    def getDescription(self):
        return self.description

    def getAmount(self):
        return self.amount

class IncomeRecord:
    def __init__(self, cellrange: CellRow):
        self.cellrange = cellrange

    def write(self, income: Income):
        recordIterator = iter(self.cellrange)
        next(recordIterator).String = income.getDescription()
        valueCell = next(recordIterator)
        valueCell.Value = income.getAmount()
        valueCell.NumberFormat = NumberFormat.CURRENCY

###############################################################################
