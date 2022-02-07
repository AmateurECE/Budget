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
from datetime import datetime

from ..account import Account, AccountHistorySummary, \
    AccountHistorySummaryRecord
from ..cellrange import CellMatrix, CellMatrixIterator
from ..cellname import getCellNameFromCoordinates, COLUMN_MAX, ROW_MAX
from ..expense import BudgetedExpense, BudgetedExpenseRecord
from ..income import Income, IncomeRecord
from ..sheet import clearSheet

class ExpenseSubForm:
    def __init__(self, iterator: CellMatrix):
        self.rowIterator = iterator

    def read(self) -> Dict[str, List[BudgetedExpense]]:
        pass

    def write(self, expenses: Dict[str, List[BudgetedExpense]]):
        expensesTitle = iter(next(self.rowIterator))
        next(expensesTitle)
        for header in ["Budgeted", "Spent", "Remaining"]:
            expensesTitleCell = next(expensesTitle)
            expensesTitleCell.String = header
            expensesTitleCell.CharWeight = BOLD
        for section in expenses:
            sectionTitleCell = next(self.rowIterator).getItem(0)
            sectionTitleCell.String = section
            sectionTitleCell.CharWeight = BOLD
            for expense in expenses[section]:
                BudgetedExpenseRecord(next(self.rowIterator)).write(expense)
            next(self.rowIterator)
        return self.rowIterator

class IncomeSubForm:
    def __init__(self, iterator: CellMatrixIterator):
        self.rowIterator = iterator

    def read(self) -> List[Income]:
        raise NotImplementedError()

    def write(self, incomes: List[Income]):
        sectionTitleCell = next(self.rowIterator).getItem(0)
        sectionTitleCell.String = "Incomes"
        sectionTitleCell.CharWeight = BOLD
        for income in incomes:
            IncomeRecord(next(self.rowIterator)).write(income)
        return self.rowIterator

class AccountHistorySummarySubForm:
    """Deals with persistence of a list of AccountHistorySummary instances"""
    def __init__(self, iterator: CellMatrixIterator):
        self.rowIterator = iterator

    def read(self) -> (List[AccountHistorySummary], datetime, datetime):
        raise NotImplementedError()

    def write(self, summaries: List[AccountHistorySummary]):
        headerRow = iter(next(self.rowIterator))
        headers = ["Account", "Starting Balance", "Current Balance",
                   "Expected Period End Balance"]
        for header in headers:
            headerCell = next(headerRow)
            headerCell.CharWeight = BOLD
            headerCell.String = header

        for summary in summaries:
            AccountHistorySummaryRecord(next(self.rowIterator)).write(summary)
        return self.rowIterator

class MonthlyBudget:
    def __init__(self, expenses: Dict[str, List[BudgetedExpense]],
                 incomes: List[Income], accounts: List[Account]):
        self.expenses = expenses
        self.incomes = incomes
        self.accounts = accounts

    def getExpenseSections(self) -> Dict[str, List[BudgetedExpense]]:
        return self.expenses

    def getIncomes(self) -> List[Income]:
        return self.incomes

    def getAccounts(self) -> List[Account]:
        return self.accounts

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
                        expense, float(defaults[section][expense])))
        return MonthlyBudget(expenses, incomes, [])

class MonthlyBudgetSheet:
    """Monthly budget form"""
    def __init__(self, sheet):
        self.sheet = sheet
        cellspec = f'A1:{getCellNameFromCoordinates(COLUMN_MAX, ROW_MAX)}'
        self.cellrange = CellMatrix(cellspec, sheet)

    def write(self, budget: MonthlyBudget):
        clearSheet(self.sheet)
        rowIterator = iter(self.cellrange)
        rowIterator = ExpenseSubForm(rowIterator).write(
            budget.getExpenseSections())
        next(rowIterator)
        rowIterator = IncomeSubForm(rowIterator).write(budget.getIncomes())
        next(rowIterator)
        AccountHistorySummarySubForm(rowIterator).write(budget.getAccounts())

    def read(self) -> MonthlyBudget:
        raise NotImplementedError()

###############################################################################
