# -*- coding: UTF-8 -*-
#! python3

# Isogeo
import isogeo_pysdk

# PyQt
from PyQt5 import Qt, QtWidgets

# QDarkStyle
import qdarkstyle

# OpenPyXl
import openpyxl

# compiled UI to load
from .ui_credits import Ui_dlg_credits


class Credits(QtWidgets.QDialog, Ui_dlg_credits):
    def __init__(self, parent=None):
        QtWidgets.QDialog.__init__(self, parent=parent)
        self.dlg_credits = Ui_dlg_credits()
        self.dlg_credits.setupUi(self)
        self.initUI()

    def initUI(self):
        """Start UI display and widgets signals and slots.
        """
        # set dependencies versions
        self.dlg_credits.lbl_dep_sdk_val.setText(isogeo_pysdk.__version__)
        self.dlg_credits.lbl_dep_xl_val.setText(openpyxl.__version__)
        self.dlg_credits.lbl_dep_qt_style_val.setText(qdarkstyle.__version__)
        self.dlg_credits.lbl_dep_qt5_val.setText(Qt.PYQT_VERSION_STR)
