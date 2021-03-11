# -*- coding: UTF-8 -*-
#! python3  # noqa: E265

# PyQt
from PyQt5 import QtWidgets

# compiled UI to load
from .ui_authentication import Ui_dlg_authentication


class Auth(QtWidgets.QDialog, Ui_dlg_authentication):
    def __init__(self, parent=None):
        QtWidgets.QDialog.__init__(self, parent=parent)
        self.setupUi(self)
