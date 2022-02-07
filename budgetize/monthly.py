###############################################################################
# NAME:             monthly.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Monthly budgetizer form
#
# CREATED:          02/06/2022
#
# LAST EDITED:      02/06/2022
###

from configparser import ConfigParser
import uno
from com.sun.star.awt.FontWeight import BOLD
from typing import Dict, List

from .cellrange import CellMatrix
from .cellname import getCellNameFromCoordinates, COLUMN_MAX, ROW_MAX
from .expense import BudgetedExpense, BudgetedExpenseRecord
from .income import Income, IncomeRecord
from .sheet import clearSheet

class MonthlyBudget:
    def __init__(self, expenses: Dict[str, List[BudgetedExpense]],
                 incomes: List[Income]):
        self.expenses = expenses
        self.incomes = incomes

    def getExpenseSections(self) -> Dict[str, List[BudgetedExpense]]:
        return self.expenses

    def getIncomes(self) -> List[Income]:
        return self.incomes

    @staticmethod
    def defaults(defaults: ConfigParser):
        expenses = {}
        incomes = []
        for section in defaults.sections():
            if "Incomes" == section:
                for income in defaults[section]:
                    incomes.append(Income(income, defaults[section][income]))
            else:
                expenses[section] = []
                for expense in defaults[section]:
                    expenses[section].append(BudgetedExpense(
                        expense, defaults[section][expense]))
        return MonthlyBudget(expenses, incomes)

class MonthlyBudgetSheet:
    """Monthly budget form"""
    def __init__(self, sheet):
        self.sheet = sheet

    def write(self, budget: MonthlyBudget):
        clearSheet(self.sheet)
        cellspec = f'A1:{getCellNameFromCoordinates(COLUMN_MAX, ROW_MAX)}'
        cellrange = CellMatrix(cellspec, self.sheet)
        rowIterator = iter(cellrange)
        expenses = budget.getExpenseSections()
        for section in expenses:
            sectionTitleCell = next(rowIterator).getItem(0)
            sectionTitleCell.String = section
            sectionTitleCell.CharWeight = BOLD
            for expense in expenses[section]:
                BudgetedExpenseRecord(next(rowIterator)).write(expense)
            next(rowIterator)

        sectionTitleCell = next(rowIterator).getItem(0)
        sectionTitleCell.String = "Incomes"
        sectionTitleCell.CharWeight = BOLD
        incomes = budget.getIncomes()
        for income in incomes:
            IncomeRecord(next(rowIterator)).write(income)

    def read(self) -> MonthlyBudget:
        raise NotImplementedError()

###############################################################################
