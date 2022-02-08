###############################################################################
# NAME:             forms.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Monthly budgetizer form
#
# CREATED:          02/06/2022
#
# LAST EDITED:      02/07/2022
###

from configparser import ConfigParser
import uno
from com.sun.star.awt.FontWeight import BOLD
from typing import Dict, List
from datetime import datetime

from ..account import AccountHistorySummary, AccountHistorySummaryRecord
from ..cellformat import NumberFormat
from ..cellrange import CellMatrix, CellMatrixIterator
from ..cellname import getCellNameFromCoordinates, COLUMN_MAX, ROW_MAX
from ..expense import BudgetedExpense, BudgetedExpenseRecord
from ..fund import SinkingFund, SinkingFundRecord
from ..income import Income, IncomeRecord
from ..sheet import clearSheet

from .expense import MonthlyExpense
from .budget import MonthlyBudget

class ExpenseSubForm:
    def __init__(self, iterator: CellMatrixIterator):
        self.rowIterator = iterator

    def read(self) -> (CellMatrixIterator, Dict[str, List[BudgetedExpense]]):
        next(self.rowIterator) # Eat the header row
        next(self.rowIterator) # Eat the total row
        currentRow = next(self.rowIterator)
        expensesBySection = {}
        while currentRow.getItem(4).String:
            sectionName = currentRow.getItem(0).String
            currentRow = next(self.rowIterator)
            expenses = []
            while currentRow.getItem(0).String:
                expenses.append(BudgetedExpenseRecord(currentRow).read())
                currentRow = next(self.rowIterator)
            if sectionName:
                expensesBySection[sectionName] = expenses
            currentRow = next(self.rowIterator)
        return (expensesBySection, self.rowIterator)

    def writeHeaders(self, iterator: CellMatrixIterator):
        titleRow = iter(next(iterator))
        secondTitleRow = iter(next(iterator))
        for header in ["Expenses", "Account", "Budgeted", "Spent",
                       "Remaining"]:
            expensesTitleCell = next(titleRow)
            expensesTitleCell.CharWeight = BOLD
            expensesTitleCell.String = header
        totalCell = next(secondTitleRow)
        totalCell.String = "Total"
        totalCell.CharWeight = BOLD
        next(secondTitleRow)
        return secondTitleRow

    def writeTotals(self, iterator, totals: List[float]):
        for total in totals:
            totalCell = next(iterator)
            totalCell.CharWeight = BOLD
            totalCell.NumberFormat = NumberFormat.CURRENCY
            totalCell.Value = total

    def write(self, expenses: Dict[str, List[BudgetedExpense]]):
        totalTitleRowIter = self.writeHeaders(self.rowIterator)
        totalBudgeted = 0.0
        totalSpent = 0.0
        for section in expenses:
            sectionHeaderIter = iter(next(self.rowIterator))
            sectionTitleCell = next(sectionHeaderIter)
            sectionTitleCell.CharWeight = BOLD
            sectionTitleCell.String = section

            budgeted = 0.0
            spent = 0.0
            for expense in expenses[section]:
                BudgetedExpenseRecord(next(self.rowIterator)).write(expense)
                budgeted += expense.getBudgeted()
                spent += expense.getSpent()

            totalBudgeted += budgeted
            totalSpent += spent
            next(sectionHeaderIter)
            self.writeTotals(
                sectionHeaderIter, [budgeted, spent, budgeted - spent])
            next(self.rowIterator)
        self.writeTotals(totalTitleRowIter, [totalBudgeted, totalSpent,
                                             totalBudgeted - totalSpent])
        return self.rowIterator

class IncomeSubForm:
    def __init__(self, iterator: CellMatrixIterator):
        self.rowIterator = iterator

    def read(self) -> List[Income]:
        next(self.rowIterator) # Eat the header row
        currentRow = next(self.rowIterator)
        incomes = []
        while currentRow.getItem(0).String:
            incomes.append(IncomeRecord(currentRow).read())
            currentRow = next(self.rowIterator)
        return incomes, self.rowIterator

    def writeHeaders(self, iterator):
        for header in ["Incomes", "", "Expected", "Received"]:
            sectionTitleCell = next(iterator)
            sectionTitleCell.String = header
            sectionTitleCell.CharWeight = BOLD

    def write(self, incomes: List[Income]):
        iterator = self.writeHeaders(iter(next(self.rowIterator)))
        totalRowIterator = iter(next(self.rowIterator))
        totalCell = next(totalRowIterator)
        totalCell.String = "Total"
        totalCell.CharWeight = BOLD

        totalExpected = 0.0
        totalReceived = 0.0
        for income in incomes:
            IncomeRecord(next(self.rowIterator)).write(income)
            totalExpected += income.getAmount()
            totalReceived += income.getReceived()

        next(totalRowIterator)
        totalCell = next(totalRowIterator)
        totalCell.CharWeight = BOLD
        totalCell.NumberFormat = NumberFormat.CURRENCY
        totalCell.Value = totalExpected

        totalCell = next(totalRowIterator)
        totalCell.CharWeight = BOLD
        totalCell.NumberFormat = NumberFormat.CURRENCY
        totalCell.Value = totalReceived
        return self.rowIterator

class AccountHistorySummarySubForm:
    """Deals with persistence of a list of AccountHistorySummary instances"""
    def __init__(self, iterator: CellMatrixIterator):
        self.rowIterator = iterator

    def read(self) -> (List[AccountHistorySummary], datetime, datetime):
        currentRow = next(self.rowIterator)
        accounts = []
        while currentRow.getItem(0).String:
            accounts.append(AccountHistorySummaryRecord(currentRow).read())
            currentRow = next(self.rowIterator)
        return accounts, self.rowIterator

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

class SinkingFundSubForm:
    def __init__(self, iterator: CellMatrixIterator):
        self.rowIterator = iterator

    def read(self) -> List[SinkingFund]:
        currentRow = next(self.rowIterator)
        funds = []
        while currentRow.getItem(0).String:
            funds.append(SinkingFundRecord(currentRow).read())
            currentRow = next(self.rowIterator)
        return funds, self.rowIterator

    def write(self, funds: List[SinkingFund]):
        headerRow = iter(next(self.rowIterator))
        headers = ["Sinking Fund", "Account", "Starting Balance",
                   "Current Balance", "Expected Period End Balance"]
        for header in headers:
            headerCell = next(headerRow)
            headerCell.CharWeight = BOLD
            headerCell.String = header

        for fund in funds:
            SinkingFundRecord(next(self.rowIterator)).write(fund)
        return self.rowIterator

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
        AccountHistorySummarySubForm(rowIterator).write(
            budget.getAccountSummaries())
        next(rowIterator)
        SinkingFundSubForm(rowIterator).write(budget.getSinkingFunds())

    def read(self) -> MonthlyBudget:
        rowIterator = iter(self.cellrange)
        expenses, rowIterator = ExpenseSubForm(rowIterator).read()
        next(rowIterator)
        incomes, rowIterator = IncomeSubForm(rowIterator).read()
        next(rowIterator)
        accounts, rowIterator = AccountHistorySummarySubForm(
            rowIterator).read()
        next(rowIterator)
        funds, _ = SinkingFundSubForm(rowIterator).read()
        return MonthlyBudget(expenses, incomes, accounts, funds)

###############################################################################
