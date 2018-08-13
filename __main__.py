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

from PyQt5 import QtCore, QtGui, QtWidgets
import qdarkstyle

# modules - UI
from modules.ui.export.ui_export import Ui_TabWidget

# modules - functional
from modules import Isogeo2xlsx
from modules import Isogeo2docx
from modules import IsogeoStats
from modules import isogeo2office_utils


if __name__ == "__main__":
    import sys
    # create the application and the main window
    app = QtWidgets.QApplication(sys.argv)
    TabWidget = QtWidgets.QTabWidget()
    ui = Ui_TabWidget()
    ui.setupUi(TabWidget)
    # setup stylesheet
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    # run
    TabWidget.show()
    sys.exit(app.exec_())
