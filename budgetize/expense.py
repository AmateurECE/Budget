###############################################################################
# NAME:             expense.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Expense types
#
# CREATED:          02/06/2022
#
# LAST EDITED:      02/06/2022
###

from .cellformat import NumberFormat
from .cellrange import CellRow

class BudgetedExpense:
    def __init__(self, description, budgeted):
        self.description = description
        self.budgeted = budgeted

    def getDescription(self):
        return self.description

    def getBudgeted(self):
        return self.budgeted

class BudgetedExpenseRecord:
    def __init__(self, cellrange: CellRow):
        self.cellrange = cellrange

    def write(self, expense: BudgetedExpense):
        recordIterator = iter(self.cellrange)
        next(recordIterator).String = expense.getDescription()
        valueCell = next(recordIterator)
        valueCell.Value = expense.getBudgeted()
        valueCell.NumberFormat = NumberFormat.CURRENCY

    def read(self) -> BudgetedExpense:
        raise NotImplementedError()

###############################################################################
