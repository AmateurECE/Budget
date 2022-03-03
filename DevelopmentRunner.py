###############################################################################
# NAME:             DevelopmentRunner.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      This Python script can be used to run the Budgetizer
#                   locally, connected to an Office server instance. This
#                   script is necessary because changes to the package are not
#                   picked up when running the macro through the "Run Macro"
#                   dialogue in a running instance of LibreOffice.
#
# CREATED:          12/01/2021
#
# LAST EDITED:      03/02/2022
###

from datetime import datetime
import calendar
import argparse
import code
import uno
from budgetize.budgetizer import Budgetizer

def getMonthName():
    return calendar.month_name[datetime.now().month]

def main():
    """Run the Budgetizer attached to a running Office server instance"""
    parser = argparse.ArgumentParser()
    parser.add_argument('-m', '--month', required=False,
                        default=getMonthName())
    args = parser.parse_args()
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        'com.sun.star.bridge.UnoUrlResolver', localContext)
    context = resolver.resolve(
        'uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext')
    desktop = context.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.Desktop', context)
    xSheetDoc = desktop.getCurrentComponent()
    # code.interact(local=dict(globals(), **locals()))
    budgetizer = Budgetizer(xSheetDoc, args.month)
    budgetizer.budgetize()

if __name__ == '__main__':
    main()

###############################################################################
