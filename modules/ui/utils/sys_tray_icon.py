# -*- coding: UTF-8 -*-
#! python3

"""
    PyQt5 System Tray Icon

    Purpose:      Get metadatas from an Isogeo share and store it into files

    Author:       Isogeo

    Python:      3.6.x
    Created:      18/12/2015
    Updated:      22/08/2018
"""

# ALERT WIP - see: https: // evileg.com/en/post/68/

# #############################################################################
# ########## Libraries #############
# ##################################

# standard library
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QGridLayout, QWidget, QCheckBox, QSystemTrayIcon, \
    QSpacerItem, QSizePolicy, QMenu, QAction, QStyle, qApp
from PyQt5.QtCore import QSize
    
    
class SysTrayIcon(QSystemTrayIcon):
    """
            Сheckbox and system tray icons.
            Will initialize in the constructor.
    """
    icon = self.tray_icon.setIcon(QMainWindow.standardIcon(QStyle.SP_ComputerIcon))
    
    # Override the class constructor
    def __init__(self):
        # Be sure to call the super class method
        self.setIcon(self.icon)
    
        '''
            Define and add steps to work with the system tray icon
            show - show window
            hide - hide window
            exit - exit from application
        '''
        show_action = QAction("Show", self)
        quit_action = QAction("Exit", self)
        hide_action = QAction("Hide", self)
        show_action.triggered.connect(self.show)
        hide_action.triggered.connect(self.hide)
        quit_action.triggered.connect(qApp.quit)
        tray_menu = QMenu()
        tray_menu.addAction(show_action)
        tray_menu.addAction(hide_action)
        tray_menu.addAction(quit_action)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
    
    # Override closeEvent, to intercept the window closing event
    # The window will be closed only if there is no check mark in the check box
    def closeEvent(self, event):
        if self.check_box.isChecked():
            event.ignore()
            self.hide()
            self.tray_icon.showMessage(
                "Tray Program",
                "Application was minimized to Tray",
                QSystemTrayIcon.Information,
                2000
            )


class SysTrayMenu(QMenu):
    def __init__(self):
        print("youhou")

    
if __name__ == "__main__":
    import sys
    
    app = QApplication(sys.argv)
    mw = MainWindow()
    mw.show()
    sys.exit(app.exec())
