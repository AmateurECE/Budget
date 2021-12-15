###############################################################################
# NAME:             Budgetizer.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Budgetizer entrypoint
#
# CREATED:          12/01/2021
#
# LAST EDITED:      12/15/2021
###

from .burndown import BurndownCalculator
from .cellrange import CellMatrix
from .cellname import ROW_MAX
from .sheet import SheetTable
from .transaction import TransactionLedger

class Budgetizer:
    def __init__(self, xSheetDoc):
        self.sheetDoc = xSheetDoc

    def runBurndown(self, transactions, beginningBalances, burndownConfig,
                    startDate):
        burndownTableSheetName = "Burndown Table"
        try:
            burndownTableSheet = self.sheetDoc.getSheets().getByName(
                burndownTableSheetName)
        except Exception:
            # Must not currently be a burndown table sheet
            burndownTableSheet = self.sheetDoc.createInstance(
                'com.sun.star.sheet.Spreadsheet')
            self.sheetDoc.getSheets().insertByName(
                burndownTableSheetName, burndownTableSheet)
        burndownCalculator = BurndownCalculator(
            transactions, beginningBalances, burndownTableSheet,
            burndownConfig)
        burndownCalculator.run(startDate)

    @staticmethod
    def getConfiguration(frontSheet, configName, tableWidth):
        for row in CellMatrix(f'A1:A{ROW_MAX}', frontSheet):
            if row.getItem(0).String == configName:
                return SheetTable(f'A{row.getRow() + 1}', tableWidth,
                                  frontSheet)
        return None

    def budgetize(self):
        nonRecurringForm = SheetTable(
            "A1", 4, self.sheetDoc.getSheets().getByName("Non Recurring"))
        recurringForm = SheetTable(
            "A1", 4, self.sheetDoc.getSheets().getByName("Recurring"))
        beginningBalances = SheetTable(
            "A1", 3, self.sheetDoc.getSheets().getByName("Balances"))
        frontSheet = self.sheetDoc.getSheets().getByName("Front")
        burndownConfig = Budgetizer.getConfiguration(
            frontSheet, 'Burndown', BurndownCalculator.MAX_CONFIG_COLUMNS)

        ledger = TransactionLedger()
        ledger.prepareNonRecurring(nonRecurringForm)
        startDate = beginningBalances.getHeaders()[1]
        endDate = beginningBalances.getHeaders()[2]
        ledger.prepareRecurring(recurringForm, startDate, endDate)

        self.runBurndown(ledger.getTransactions(), beginningBalances,
                         burndownConfig, startDate)

###############################################################################
