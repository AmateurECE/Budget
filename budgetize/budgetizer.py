###############################################################################
# NAME:             Budgetizer.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Budgetizer entrypoint
#
# CREATED:          12/01/2021
#
# LAST EDITED:      12/03/2021
###

from .burndown import BurndownCalculator
from .sheet import SheetTable

class Budgetizer:
    def __init__(self, xSheetDoc):
        self.sheetDoc = xSheetDoc

    def runBurndown(self, nonRecurringForm, beginningBalances):
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
            nonRecurringForm, beginningBalances, burndownTableSheet)
        burndownCalculator.run()

    def budgetize(self):
        nonRecurringForm = SheetTable(
            "A1", 4, self.sheetDoc.getSheets().getByName("Non Recurring"))
        beginningBalances = SheetTable(
            "A1", 3, self.sheetDoc.getSheets().getByName("Balances"))
        self.runBurndown(nonRecurringForm, beginningBalances)

###############################################################################
