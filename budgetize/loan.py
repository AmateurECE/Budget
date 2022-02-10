###############################################################################
# NAME:             loan.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Logic to handle loans
#
# CREATED:          02/07/2022
#
# LAST EDITED:      02/08/2022
###

from .cellrange import CellRow
from .cellformat import NumberFormat

import math

class Loan:
    def __init__(self, name, interest, startingBalance):
        self.name = name
        self.interest = interest
        self.startingBalance = startingBalance
        self.endingBalance = startingBalance

    def getName(self):
        return self.name

    def getStartingBalance(self):
        return self.startingBalance

    def getAPR(self):
        return self.interest

    def setEndingBalance(self, endingBalance):
        self.endingBalance = endingBalance

    def getEndingBalance(self):
        return self.endingBalance

    def getPayoffPeriod(self):
        payment = self.startingBalance - self.endingBalance
        time = (
            math.log(1 - (self.interest / 12 * self.endingBalance / payment))
            / math.log(1 / (1 + (self.interest / 12))))
        return round(time)

class LoanRecord:
    def __init__(self, cellrange: CellRow):
        self.cellrange = iter(cellrange)

    def read(self) -> Loan:
        name = next(self.cellrange).String
        apr = next(self.cellrange).Value
        startingBalance = next(self.cellrange).Value
        return Loan(name, apr, startingBalance)

    def write(self, loan: Loan):
        next(self.cellrange).String = loan.getName()
        next(self.cellrange).Value = loan.getAPR()
        startingBalanceCell = next(self.cellrange)
        startingBalanceCell.NumberFormat = NumberFormat.CURRENCY
        startingBalanceCell.Value = loan.getStartingBalance()

        endingBalanceCell = next(self.cellrange)
        endingBalanceCell.NumberFormat = NumberFormat.CURRENCY
        endingBalanceCell.Value = loan.getEndingBalance()
        next(self.cellrange).String = f'{loan.getPayoffPeriod()} mo'

###############################################################################
