###############################################################################
# NAME:             budgetize.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Perform some calculations on an input spreadsheet.
#
# CREATED:          11/30/2021
#
# LAST EDITED:      11/30/2021
###

import toml

def testBudgetize():
    xSheet = XSCRIPTCONTEXT.getDocument().getCurrentController()
    xSheet.getSelection().getCellByPosition(0, 0).setString('Budgetize!')

# Lists the scripts, that shall be visible inside OOo. Can be omitted, if all
# functions shall be visible.
g_exportedScripts = testBudgetize,

###############################################################################
