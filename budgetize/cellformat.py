###############################################################################
# NAME:             cellformat.py
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Common logic for formatting cells
#
# CREATED:          12/03/2021
#
# LAST EDITED:      02/06/2022
###

from com.sun.star.util import NumberFormat

class NumberFormat:
    CURRENCY = NumberFormat.CURRENCY | NumberFormat.SCIENTIFIC \
        | NumberFormat.FRACTION

###############################################################################
