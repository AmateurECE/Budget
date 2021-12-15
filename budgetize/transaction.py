###############################################################################
# NAME:             transaction.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Logic encapsulating account transactions
#
# CREATED:          12/14/2021
#
# LAST EDITED:      12/15/2021
###

from datetime import datetime
from pyrecurrence import PyOccurrenceSeries

class Transaction:
    def __init__(self, description='', amount=0.0, accountName='', date=None):
        self.description = description
        self.amount = amount
        self.accountName = accountName
        self.date = date

    def applyToAccount(self, account):
        account.updateBalance(self.amount)

class RecurringTransactionIterator:
    def __init__(self, transactionTemplate=None, schedule=None, startDate=None,
                 endDate=None):
        self.transactionTemplate = transactionTemplate
        self.schedule = schedule
        self.startDate = startDate
        self.endDate = endDate

    def __next__(self):
        self.startDate = self.schedule.next_occurrence(self.startDate)
        if self.startDate > self.endDate:
            raise StopIteration()
        description = self.transactionTemplate.description
        amount = self.transactionTemplate.amount
        accountName = self.transactionTemplate.accountName
        return Transaction(description=description, amount=amount,
                           accountName=accountName, date=self.startDate)

class RecurringTransaction:
    def __init__(self, transactionTemplate=None, schedule='', startDate=None,
                 endDate=None):
        self.transactionTemplate = transactionTemplate
        self.schedule = PyOccurrenceSeries(schedule)
        self.startDate = startDate
        self.endDate = endDate

    def __iter__(self):
        return RecurringTransactionIterator(
            transactionTemplate=self.transactionTemplate,
            schedule=self.schedule, startDate=self.startDate,
            endDate=self.endDate)

class TransactionLedger:
    def __init__(self):
        self.transactions = []

    def prepareNonRecurring(self, nonRecurringForm):
        for row in nonRecurringForm:
            self.transactions.append(Transaction(
                description=row['Description'].String,
                date=datetime.strptime(row['Date'].String, '%m/%d/%y'),
                accountName=row['Account'].String, amount=row['Amount'].Value))

    def prepareRecurring(self, recurringForm, startDate, endDate):
        for row in recurringForm:
            transactionTemplate = Transaction(
                description=row['Description'].String,
                amount=row['Amount'].Value, accountName=row['Account'].String)
            recurringTransaction = RecurringTransaction(
                transactionTemplate=transactionTemplate,
                schedule=row['Schedule'].String.lower(),
                startDate=datetime.strptime(startDate, '%m/%d/%y'),
                endDate=datetime.strptime(endDate, '%m/%d/%y'))
            for transaction in recurringTransaction:
                self.transactions.append(transaction)

    def getTransactions(self):
        return sorted(self.transactions, key=lambda x: x.date)

###############################################################################
