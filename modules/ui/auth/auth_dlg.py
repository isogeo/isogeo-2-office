# -*- coding: UTF-8 -*-
#! python3

# PyQt
from PyQt5.QtWidgets import QDialog

# compiled UI to load
from .ui_authentication import Ui_dlg_authentication


class Auth(QDialog, Ui_dlg_authentication):
    def __init__(self, parent=None):
        QDialog.__init__(self, parent=parent)
        self.setupUi(self)
