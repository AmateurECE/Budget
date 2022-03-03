###############################################################################
# NAME:             Budgetizer.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Budgetizer entrypoint
#
# CREATED:          12/01/2021
#
# LAST EDITED:      03/02/2022
###

from configparser import ConfigParser
from appdirs import user_data_dir
import uno
from com.sun.star.container import NoSuchElementException

from .monthly.budget import MonthlyBudget
from .monthly.forms import MonthlyBudgetSheet
from .monthly.expense import MonthlyExpenseSheet, MonthlyExpense

APP_NAME = 'Budgetizer'
APP_AUTHOR = "edtwardy"

class Budgetizer:
    def __init__(self, xSheetDoc, month):
        self.sheetDoc = xSheetDoc
        self.month = month

    def initBudgetSheet(self):
        sheetName = self.month + ' Budget'
        try:
            monthlySheet = self.sheetDoc.getSheets().getByName(sheetName)
            monthlyBudget = MonthlyBudgetSheet(monthlySheet).read()
        except NoSuchElementException:
            monthlySheet = self.sheetDoc.createInstance(
                'com.sun.star.sheet.Spreadsheet')
            self.sheetDoc.getSheets().insertByName(sheetName, monthlySheet)
            config = ConfigParser()
            config.optionxform=str
            config.read(user_data_dir(APP_NAME, APP_AUTHOR) + '/defaults.ini')
            monthlyBudget = MonthlyBudget.defaults(config)
        return (monthlyBudget, monthlySheet)

    def initExpensesSheet(self):
        sheetName = self.month + ' Expenses'
        try:
            expenseSheet = self.sheetDoc.getSheets().getByName(sheetName)
            monthlyExpenses = MonthlyExpenseSheet(expenseSheet).read()
        except NoSuchElementException:
            expenseSheet = self.sheetDoc.createInstance(
                'com.sun.star.sheet.Spreadsheet')
            self.sheetDoc.getSheets().insertByName(sheetName, expenseSheet)
            monthlyExpenses = []
            MonthlyExpenseSheet(expenseSheet).write(monthlyExpenses)
        return monthlyExpenses

    def budgetize(self):
        (budget, budgetSheet) = self.initBudgetSheet()
        expenses = self.initExpensesSheet()
        budget.calculateExpectedBalances()
        budget.applyExpenses(expenses)
        MonthlyBudgetSheet(budgetSheet).write(budget)

###############################################################################
