###############################################################################
# NAME:             budget.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Budget logic
#
# CREATED:          02/07/2022
#
# LAST EDITED:      02/08/2022
###

from configparser import ConfigParser
from datetime import datetime
from typing import Dict, List
import re

from ..fund import SinkingFund
from ..account import AccountHistorySummary
from ..income import Income
from ..expense import BudgetedExpense
from .expense import MonthlyExpense
from ..loan import Loan

class MonthlyBudget:
    def __init__(self, expenses: Dict[str, List[BudgetedExpense]],
                 incomes: List[Income],
                 accountSummaries: List[AccountHistorySummary],
                 funds: List[SinkingFund], loans: List[Loan]):
        self.expenses = expenses
        self.incomes = incomes
        self.accountSummaries = accountSummaries
        self.funds = funds
        self.loans = loans

    def getExpenseSections(self) -> Dict[str, List[BudgetedExpense]]:
        return self.expenses

    def getIncomes(self) -> List[Income]:
        return self.incomes

    def getAccountSummaries(self) -> List[AccountHistorySummary]:
        return self.accountSummaries

    def getSinkingFunds(self) -> List[SinkingFund]:
        return self.funds

    def getLoans(self) -> List[Loan]:
        return self.loans

    def getBudgetedIncome(self, income: MonthlyExpense):
        income = next(filter(
            lambda x: x.getDescription() == income.getDescription(),
            self.incomes), None)
        if not income:
            raise RuntimeError(
                f'The transaction {income.getDescription()} was not an '
                + 'expected income'
            )
        return income

    def getBudgetedExpense(self, expense: MonthlyExpense):
        try:
            category = next(filter(lambda x: x == expense.getCategory(),
                                   self.expenses.keys()), None)
            if not category:
                raise KeyError(expense.getCategory())
        except KeyError:
            raise RuntimeError(f'The category {expense.getCategory()} is not'
                               + ' currently accounted for in the budget!')
        try:
            budgeted = next(filter(
                lambda x: x.getDescription() == expense.getLineItem(),
                self.expenses[category]), None)
            if not budgeted:
                raise KeyError(expense.getLineItem())
            return budgeted
        except KeyError:
            raise RuntimeError(f'The line item {expense.getLineItem()} is not'
                               + ' currently accounted for in the budget!')

    def getAccountByName(self, accountName) -> AccountHistorySummary:
        account = next(filter(lambda x: x.getAccountName() == accountName,
                              self.accountSummaries), None)
        if not account:
            raise RuntimeError(f'No account named {accountName}')
        return account

    def getLoanByName(self, loanName) -> Loan:
        loan = next(filter(lambda x: x.getName() == loanName, self.loans),
                    None)
        if not loan:
            raise RuntimeError(f'No loan named {loanName}')
        return loan

    def ensureAccountsForExpense(self, budgetedExpense, expense):
        if budgetedExpense.getAccountName() != expense.getAccountName():
            raise RuntimeError(
                f'{budgetedExpense.getDescription()} should come out of '
                + f'{budgetedExpense.getAccountName()}, but it actually '
                + f'came out of {expense.getAccountName()}'
            )

    def ensureAccountsForIncome(self, budgetedIncome, income):
        if budgetedIncome.getAccountName() != income.getAccountName():
            raise RuntimeError(
                f'{budgetedIncome.getDescription()} should go into '
                + f'{budgetedIncome.getAccountName()}, but it actually '
                + f'came went into {income.getAccountName()}'
            )

    def applyIncomeTransaction(self, income: MonthlyExpense):
        budgetedIncome = self.getBudgetedIncome(income)
        self.ensureAccountsForIncome(budgetedIncome, income)
        account = self.getAccountByName(budgetedIncome.getAccountName())
        budgetedIncome.receive(income.getAmount())
        account.updateBalance(income.getAmount())

    def applyExpenseTransaction(self, expense: MonthlyExpense):
        budgetedExpense = self.getBudgetedExpense(expense)
        self.ensureAccountsForExpense(budgetedExpense, expense)
        account = self.getAccountByName(budgetedExpense.getAccountName())
        budgetedExpense.spend(expense.getAmount())
        account.updateBalance(expense.getAmount())

    def applyExpenses(self, expenses: List[MonthlyExpense]):
        now = datetime.now()
        for expense in expenses:
            if expense.getDate() > now:
                continue # Skip expenses that haven't happened yet.

            if expense.getAmount() < 0:
                self.applyExpenseTransaction(expense)
            else:
                self.applyIncomeTransaction(expense)

    def updateExpectedBalanceForExpenseAccounts(self, expense):
        matches = re.fullmatch(r'transfer\(([^,]*),([^,]*)\)',
                               expense.getAccountName())
        if matches:
            account = self.getAccountByName(matches.group(1))
            loan = self.getLoanByName(matches.group(2))
            loan.setEndingBalance(
                loan.getEndingBalance() - expense.getBudgeted())
            account.updateExpectedBalance(-1 * expense.getBudgeted())
        else:
            self.getAccountByName(
                expense.getAccountName()).updateExpectedBalance(
                    -1 * expense.getBudgeted())

    def calculateExpectedBalances(self):
        if not self.accountSummaries:
            return # Budget generated by "defaults" has no accounts

        for income in self.incomes:
            self.getAccountByName(
                income.getAccountName()).updateExpectedBalance(
                    income.getAmount())

        for _, category in self.expenses.items():
            for expense in category:
                self.updateExpectedBalanceForExpenseAccounts(expense)

    @staticmethod
    def defaults(defaults: ConfigParser):
        expenses = {}
        incomes = []
        for section in defaults.sections():
            if "Incomes" == section:
                for income in defaults[section]:
                    incomes.append(Income(
                        income, "", float(defaults[section][income])))
            else:
                expenses[section] = []
                for expense in defaults[section]:
                    expenses[section].append(BudgetedExpense(
                        expense, "", float(defaults[section][expense]), 0.0))
        return MonthlyBudget(expenses, incomes, [], [], [])

###############################################################################
