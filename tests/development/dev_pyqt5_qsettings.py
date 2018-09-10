# -*- coding: UTF-8 -*-
#! python3

import sys
from PyQt5.QtCore import QSettings
from PyQt5.QtWidgets import QApplication

# #############################################################################
# ##### Functions ##################
# ##################################

def settings_cleaner(settings_class=QSettings, group_to_remove=""):
    """Clear stored QSettings. By default remove everything.

    See: http://doc.qt.io/qt-5/qsettings.html#clear
      and http://doc.qt.io/qt-5/qsettings.html#remove
    """
    settings_class.remove(group_to_remove)



# #############################################################################
# ##### Main #######################
# ##################################

app = QApplication(sys.argv)

app.settings = QSettings('Isogeo', 'IsogeoToOffice')

print(app.settings.childGroups())
print(app.settings.allKeys())

settings_cleaner(settings_class=app.settings,
                 group_to_remove="log")

print(app.settings.childGroups())

#sys.exit(app.exec_())
