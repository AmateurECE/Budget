#!/bin/sh
###############################################################################
# NAME:             start.sh
#
# AUTHOR:           Ethan D. Twardy <ethan.twardy@gmail.com>
#
# DESCRIPTION:      Start up the LibreOffice server
#
# CREATED:          02/03/2022
#
# LAST EDITED:      02/03/2022
###

soffice --calc \
    --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager" &

###############################################################################
