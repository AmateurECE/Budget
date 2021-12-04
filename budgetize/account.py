###############################################################################
# NAME:             account.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Account classes
#
# CREATED:          12/03/2021
#
# LAST EDITED:      12/03/2021
###

class Account:
    def __init__(self, name, balance):
        self.name = name
        self.balance = balance

    def updateBalance(self, diff):
        self.balance += diff

    def getName(self):
        return self.name

    def getBalance(self):
        return self.balance

###############################################################################
