# -*- coding: UTF-8 -*-
#! python3

"""
    Isogeo To Office - SystemTray Menu

    Author:      Isogeo
    Python:      3.6.x
"""

# #############################################################################
# ########## Libraries #############
# ##################################

# standard library
import logging

# 3rd party
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QAction

# #############################################################################
# ########## Classes ###############
# ##################################


class SystrayMenu(QSystemTrayIcon):
    """
        Isogeo to Office  SystemTray Menu derived from QSystemTrayIcon
    """

    def __init__(self, parent=None, caption="Isogeo To Office"):
        QSystemTrayIcon.__init__(self, parent)
        # tooltip
        self.setToolTip(caption)
        # set menu
        self.menu = QMenu(parent)
        # Available actions in Systray menu
        self.act_show = QAction(
            QIcon("resources/systray/window-restore.svg"), self.tr("Show"), self
        )
        self.act_hide = QAction(
            QIcon("resources/systray/window-minimize.svg"), self.tr("Hide"), self
        )
        self.act_quit = QAction(
            QIcon("resources/systray/window-close-o.svg"), self.tr("Exit"), self
        )
        # add actions to the menu
        self.menu.addActions([self.act_show, self.act_hide, self.act_quit])
        self.setContextMenu(self.menu)


# #############################################################################
if __name__ == "__main__":
    logging.warning("Meant to be called by other program")
