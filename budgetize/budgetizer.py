###############################################################################
# NAME:             Budgetizer.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Budgetizer entrypoint
#
# CREATED:          12/01/2021
#
# LAST EDITED:      02/03/2022
###

from .account import AccountHistorySummaryForm
from .burndown import BurndownCalculator
from .cellrange import CellMatrix
from .cellname import ROW_MAX
from .sheet import SheetTable
from .transaction import TransactionLedger, RecurringTransactionForm, \
    TransactionForm

class Budgetizer:
    def __init__(self, xSheetDoc):
        self.sheetDoc = xSheetDoc

    def getOrCreateSheet(self, sheetName):
        try:
            return self.sheetDoc.getSheets().getByName(sheetName)
        except Exception:
            # Must not currently be a burndown table sheet
            namedSheet = self.sheetDoc.createInstance(
                'com.sun.star.sheet.Spreadsheet')
            self.sheetDoc.getSheets().insertByName(namedSheet, sheetName)
            return namedSheet

    def runBurndown(self, transactions, balances: AccountHistorySummaryForm,
                    burndownConfig):
        burndownTableSheet = self.getOrCreateSheet("Burndown Table")
        burndownCalculator = BurndownCalculator(
            transactions, balances, burndownTableSheet,
            burndownConfig)
        burndownCalculator.run(balances.getStartDate(), balances.getEndDate())

    @staticmethod
    def getConfiguration(frontSheet, configName, tableWidth):
        for row in CellMatrix(f'A1:A{ROW_MAX}', frontSheet):
            if row.getItem(0).String == configName:
                return SheetTable(f'A{row.getRow() + 1}', tableWidth,
                                  frontSheet)
        return None

    def budgetize(self):
        nonRecurringForm = TransactionForm(SheetTable(
            "A1", 4, self.sheetDoc.getSheets().getByName("Non Recurring")))
        recurringForm = RecurringTransactionForm(SheetTable(
            "A1", 4, self.sheetDoc.getSheets().getByName("Recurring")))
        balances = AccountHistorySummaryForm(SheetTable(
            "A1", 3, self.sheetDoc.getSheets().getByName("Balances")))
        frontSheet = self.sheetDoc.getSheets().getByName("Front")
        burndownConfig = Budgetizer.getConfiguration(
            frontSheet, 'Burndown', BurndownCalculator.MAX_CONFIG_COLUMNS)

        ledger = TransactionLedger(recurringForm, nonRecurringForm)
        startDate = balances.getStartDate()
        endDate = balances.getEndDate()

        self.runBurndown(ledger.getTransactions(startDate, endDate), balances,
                         burndownConfig)

###############################################################################
