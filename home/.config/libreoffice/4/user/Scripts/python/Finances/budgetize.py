###############################################################################
# NAME:             budgetize.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Perform some calculations on an input spreadsheet.
#
# CREATED:          11/30/2021
#
# LAST EDITED:      12/02/2021
###

from budgetize.budgetizer import Budgetizer

def runBudgetizer():
    budgetizer = Budgetizer(XSCRIPTCONTEXT.getDesktop().getCurrentComponent())
    budgetizer.budgetize()

# Lists the scripts, that shall be visible inside OOo. Can be omitted, if all
# functions shall be visible.
g_exportedScripts = runBudgetizer,

###############################################################################
