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
# LAST EDITED:      12/01/2021
###

import uno
from budgetize.Budgetizer import Budgetizer

def main():
    """Run the Budgetizer attached to a running Office server instance"""
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        'com.sun.star.bridge.UnoUrlResolver', localContext)
    context = resolver.resolve(
        'uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext')
    desktop = context.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.Desktop', context)
    xSheetDoc = desktop.getCurrentComponent()
    budgetizer = Budgetizer(xSheetDoc)
    budgetizer.budgetize()

if __name__ == '__main__':
    main()

###############################################################################
