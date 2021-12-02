###############################################################################
# NAME:             budgetize.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Perform some calculations on an input spreadsheet.
#
# CREATED:          11/30/2021
#
# LAST EDITED:      12/01/2021
###

from budgetize.Budgetizer import Budgetizer

def runBudgetizer():
    xSheetDoc = XSCRIPTCONTEXT.getDesktop().getCurrentComponent()
    budgetizer = Budgetizer(xSheetDoc.CurrentController)
    budgetizer.budgetize()

# Lists the scripts, that shall be visible inside OOo. Can be omitted, if all
# functions shall be visible.
g_exportedScripts = runBudgetizer,

###############################################################################
