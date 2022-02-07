###############################################################################
# NAME:             expense.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Logic for handling monthly expenses
#
# CREATED:          02/06/2022
#
# LAST EDITED:      02/07/2022
###

from datetime import datetime
from typing import List

import uno
from com.sun.star.awt.FontWeight import BOLD

from ..cellformat import NumberFormat
from ..cellname import getCellNameFromCoordinates, COLUMN_MAX, ROW_MAX
from ..cellrange import CellMatrix, CellRow
from ..sheet import SheetTable

class MonthlyExpense:
    def __init__(self, description, category, date, amount, accountName):
        self.description = description
        self.category = category
        self.date = date
        self.amount = amount
        self.accountName = accountName

    def getDescription(self):
        return self.description

    def getCategory(self):
        return self.category

    def getDate(self):
        return self.date

    def getAmount(self):
        return self.amount

    def getAccountName(self):
        return self.accountName

class MonthlyExpenseRecord:
    def __init__(self, cellrange: CellRow):
        self.cellrange = cellrange

    def read(self) -> MonthlyExpense:
        recordIter = iter(self.cellrange)
        description = next(recordIter).String
        category = next(recordIter).String
        date = datetime.strptime(next(recordIter).String, '%m/%d/%y')
        amount = next(recordIter).Value
        accountName = next(recordIter).String
        return MonthlyExpense(description, category, date, amount, accountName)

    def write(self, expense: MonthlyExpense):
        recordIter = iter(self.cellrange)
        next(recordIter).String = expense.getDescription()
        next(recordIter).String = expense.getCategory()
        next(recordIter).String = datetime.strftime(
            expense.getDate(), '%m/%d/%y')

        amountCell = next(recordIter)
        amountCell.Value = expense.getAmount()
        amountCell.NumberFormat = NumberFormat.CURRENCY

        next(recordIter).String = expense.getAccountName()

class MonthlyExpenseSheet:
    def __init__(self, sheet):
        self.sheet = sheet
        cellspec = f'A1:{getCellNameFromCoordinates(COLUMN_MAX, ROW_MAX)}'
        self.cellrange = CellMatrix(cellspec, sheet)

    def write(self, expenses: List[MonthlyExpense]):
        rowIter = iter(self.cellrange)
        recordIter = iter(next(rowIter))
        for header in ["Description", "Category", "Date", "Amount", "Account"]:
            headerCell = next(recordIter)
            headerCell.String = header
            headerCell.CharWeight = BOLD
        for expense in expenses:
            MonthlyExpenseRecord(next(rowIter)).write(expense)

    def read(self) -> List[MonthlyExpense]:
        sheetTable = SheetTable('A1', 5, self.sheet)
        expenses = []
        for row in sheetTable:
            expenses.append(MonthlyExpenseRecord(row).read())
        return expenses

###############################################################################
