###############################################################################
# NAME:             transaction.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Logic encapsulating account transactions
#
# CREATED:          12/14/2021
#
# LAST EDITED:      02/03/2022
###

from datetime import datetime
from typing import List

from pyrecurrence import PyOccurrenceSeries

from .cellrange import CellRow
from .sheet import SheetTable

###############################################################################
# Transactions
###

class Transaction:
    def __init__(self, description='', amount=0.0, accountName='', date=None):
        self.description = description
        self.amount = amount
        self.accountName = accountName
        self.date = date

    def applyToAccount(self, account):
        account.updateBalance(self.amount)

class TransactionRecord:
    def __init__(self, cellrange: CellRow):
        self.cellrange = cellrange

    def read(self) -> Transaction:
        iterator = iter(self.cellrange)
        date = datetime.strptime(next(iterator).String, '%m/%d/%y')
        description = next(iterator).String
        amount = next(iterator).Value
        accountName = next(iterator).String
        return Transaction(description, amount, accountName, date)

class TransactionForm:
    def __init__(self, cellrange: SheetTable):
        self.cellrange = cellrange

    def read(self) -> List[TransactionRecord]:
        result = []
        for row in self.cellrange:
            result.append(TransactionRecord(row).read())
        return result

###############################################################################
# Recurring Transactions
###

class RecurringTransactionIterator:
    """Iterate over the series of payments in a recurring transaction"""
    def __init__(self, template: Transaction=None, schedule=None,
                 startDate=None, endDate=None):
        self.template = template
        self.schedule = schedule
        self.startDate = startDate
        self.endDate = endDate

    def __next__(self):
        self.startDate = self.schedule.next_occurrence(self.startDate)
        if self.startDate > self.endDate:
            raise StopIteration()
        description = self.template.description
        amount = self.template.amount
        accountName = self.template.accountName
        return Transaction(description=description, amount=amount,
                           accountName=accountName, date=self.startDate)

class RecurringTransactionIteratorGenerator:
    """Generates an iterator for a recurring transaction and start/end dates"""
    def __init__(self, transaction, startDate, endDate):
        self.template = transaction.template
        self.schedule = transaction.schedule
        self.startDate = startDate
        self.endDate = endDate

    def __iter__(self):
        return RecurringTransactionIterator(
            template=self.template, schedule=self.schedule,
            startDate=self.startDate, endDate=self.endDate)

class RecurringTransaction:
    """Encapsulates a payment schedule from a template transaction"""
    def __init__(self, template=None, schedule=''):
        self.template = template
        self.schedule = PyOccurrenceSeries(schedule)

    def forDates(self, startDate, endDate):
        return RecurringTransactionIteratorGenerator(self, startDate, endDate)

class RecurringTransactionRecord:
    """Concerned with serialization of a single recurring transaction"""
    def __init__(self, cellrange: CellRow):
        self.cellrange = cellrange

    def read(self) -> RecurringTransaction:
        iterator = iter(self.cellrange)
        description = next(iterator).String
        amount = next(iterator).Value
        accountName = next(iterator).String
        schedule = next(iterator).String.lower()
        template = Transaction(description, amount, accountName)
        return RecurringTransaction(template, schedule)

class RecurringTransactionForm:
    """Deals with persistence of recurring transactions"""
    def __init__(self, cellrange: SheetTable):
        self.cellrange = cellrange

    def read(self) -> List[RecurringTransaction]:
        result = []
        for row in self.cellrange:
            result.append(RecurringTransactionRecord(row).read())
        return result

###############################################################################
# Transaction Ledger
###

class TransactionLedger:
    """Sort and coalesce a list of transactions based on the contents of a
    RecurringTransactionForm and a TransactionForm"""
    def __init__(self, recurringForm, nonRecurringForm):
        self.recurringForm = recurringForm
        self.nonRecurringForm = nonRecurringForm

    def getTransactions(self, startDate, endDate):
        transactions = self.nonRecurringForm.read()

        startDateObj = datetime.strptime(startDate, '%m/%d/%y')
        endDateObj = datetime.strptime(endDate, '%m/%d/%y')
        for transaction in self.recurringForm.read():
            for entry in transaction.forDates(startDateObj, endDateObj):
                transactions.append(entry)

        return sorted(transactions, key=lambda x: x.date)

###############################################################################
