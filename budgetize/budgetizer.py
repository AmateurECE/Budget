###############################################################################
# NAME:             Budgetizer.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Budgetizer entrypoint
#
# CREATED:          12/01/2021
#
# LAST EDITED:      02/06/2022
###

import calendar
from datetime import datetime
from configparser import ConfigParser

from appdirs import user_data_dir
import uno
from com.sun.star.container import NoSuchElementException

from .monthly import MonthlyBudgetSheet, MonthlyBudget

APP_NAME = 'Budgetizer'
APP_AUTHOR = "edtwardy"

def getMonthName():
    return calendar.month_name[datetime.now().month]

class Budgetizer:
    def __init__(self, xSheetDoc):
        self.sheetDoc = xSheetDoc

    def budgetize(self):
        sheetName = getMonthName() + ' Budget'
        try:
            monthlySheet = self.sheetDoc.getSheets().getByName(sheetName)
        except NoSuchElementException:
            monthlySheet = self.sheetDoc.createInstance(
                'com.sun.star.sheet.Spreadsheet')
            self.sheetDoc.getSheets().insertByName(sheetName, monthlySheet)
        # TODO: Put these under the except branch
        config = ConfigParser()
        config.optionxform=str
        config.read(user_data_dir(APP_NAME, APP_AUTHOR) + '/defaults.ini')
        monthlyBudget = MonthlyBudget.defaults(config)
        MonthlyBudgetSheet(monthlySheet).write(monthlyBudget)

###############################################################################
