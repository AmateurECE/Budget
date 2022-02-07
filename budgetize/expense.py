###############################################################################
# NAME:             expense.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Expense types
#
# CREATED:          02/06/2022
#
# LAST EDITED:      02/07/2022
###

from .cellformat import NumberFormat
from .cellrange import CellRow

class BudgetedExpense:
    def __init__(self, description, accountName, budgeted, spent):
        self.description = description
        self.accountName = accountName
        self.budgeted = budgeted
        self.spent = spent

    def getDescription(self):
        return self.description

    def getAccountName(self):
        return self.accountName

    def getBudgeted(self):
        return self.budgeted

    def getSpent(self):
        return self.spent

    def spend(self, amount):
        self.spend - amount

class BudgetedExpenseRecord:
    def __init__(self, cellrange: CellRow):
        self.cellrange = cellrange

    def write(self, expense: BudgetedExpense):
        recordIterator = iter(self.cellrange)
        next(recordIterator).String = expense.getDescription()
        next(recordIterator).String = expense.getAccountName()

        budgetedCell = next(recordIterator)
        budgetedCell.Value = expense.getBudgeted()
        budgetedCell.NumberFormat = NumberFormat.CURRENCY

        spentCell = next(recordIterator)
        spentCell.Value = expense.getSpent()
        spentCell.NumberFormat = NumberFormat.CURRENCY

        remainingCell = next(recordIterator)
        remainingCell.Value = expense.getBudgeted() - expense.getSpent()
        remainingCell.NumberFormat = NumberFormat.CURRENCY

    def read(self) -> BudgetedExpense:
        recordIterator = iter(self.cellrange)
        description = next(recordIterator).String
        accountName = next(recordIterator).String
        budgeted = next(recordIterator).Value
        spent = next(recordIterator).Value
        return BudgetedExpense(description, accountName, budgeted, spent)

###############################################################################
