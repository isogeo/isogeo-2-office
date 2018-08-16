# -*- coding: UTF-8 -*-
#! python3

# PyQt
from PyQt5.QtWidgets import QDialog

# compiled UI to load
from .ui_credits import Ui_dlg_credits

class Credits(QDialog, Ui_dlg_credits):
    def  __init__(self, parent=None):
        QDialog.__init__(self, parent=parent)
        self.setupUi(self)
