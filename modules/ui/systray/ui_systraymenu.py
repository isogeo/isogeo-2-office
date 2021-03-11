# -*- coding: UTF-8 -*-
#! python3  # noqa: E265

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
from PyQt5 import QtGui, QtWidgets

# #############################################################################
# ########## Classes ###############
# ##################################


class SystrayMenu(QtWidgets.QSystemTrayIcon):
    """
        Isogeo to Office  SystemTray Menu derived from QtWidgets.QSystemTrayIcon
    """

    def __init__(self, parent=None, caption="Isogeo To Office"):
        QtWidgets.QSystemTrayIcon.__init__(self, parent)
        # tooltip
        self.setToolTip(caption)
        # set menu
        self.menu = QtWidgets.QMenu(parent)
        # Available actions in Systray menu
        self.act_show = QtWidgets.QAction(
            QtGui.QIcon("resources/systray/window-restore.svg"), self.tr("Show"), self
        )
        self.act_hide = QtWidgets.QAction(
            QtGui.QIcon("resources/systray/window-minimize.svg"), self.tr("Hide"), self
        )
        self.act_quit = QtWidgets.QAction(
            QtGui.QIcon("resources/systray/window-close-o.svg"), self.tr("Exit"), self
        )
        # add actions to the menu
        self.menu.addActions([self.act_show, self.act_hide, self.act_quit])
        self.setContextMenu(self.menu)


# #############################################################################
if __name__ == "__main__":
    logging.warning("Meant to be called by other program")
