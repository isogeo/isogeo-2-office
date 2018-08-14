# -*- coding: UTF-8 -*-
#! python3

"""
    Isogeo To Office - Main launcher

    Purpose:      Get metadatas from an Isogeo share and store it into files

    Author:       Julien Moura (@geojulien)

     Python:      3.6.x
    Created:      18/12/2015
    Updated:      22/08/2018
"""

import qdarkstyle

# submodules - UI
from modules.ui.auth.ui_authentication import Ui_dlg_authentication
from modules.ui.credits.ui_credits import Ui_dlg_credits
from modules.ui.main.ui_IsogeoToOffice import Ui_tabs_IsogeoToOffice

# modules - functional
from modules import Isogeo2xlsx
from modules import Isogeo2docx
from modules import IsogeoStats
from modules import isogeo2office_utils


# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == "__main__":
    import sys
    # create the application and the main window
    app = QtWidgets.QApplication(sys.argv)
    # apply dark style
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    # apply language
    locale_path = path.join(app_dir,
                            'i18n',
                            'IsogeoToOffice_{}.qm'.format(current_locale.system().name()))
    translator = QtCore.QTranslator()
    translator.load(path.realpath(locale_path))
    app.installTranslator(translator)
    # link to Isogeo to Office main UI
    i2o = IsogeoToOffice_Main()
    sys.exit(app.exec_())
