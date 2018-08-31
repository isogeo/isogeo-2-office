# -*- coding: UTF-8 -*-
#! python3

"""
    Isogeo To Office - Threads used to subprocess some tasks


    Author:       Julien Moura (@geojulien)

    Python:      3.6.x
    Created:      18/12/2015
    Updated:      22/08/2018
"""

# #############################################################################
# ########## Libraries #############
# ##################################

# standard library
from datetime import datetime
import logging
from functools import partial
from logging.handlers import RotatingFileHandler
from os import path
import platform

# 3rd party library
from isogeo_pysdk import __version__ as pysdk_version
from PyQt5.QtCore import (QBasicTimer, QDate, QLocale, QSettings, QTimerEvent,
                          QTranslator, QThread, pyqtSignal, pyqtSlot)
from PyQt5.QtGui import QCloseEvent
from PyQt5.QtWidgets import (QApplication, QComboBox, QDialog,
                             QMessageBox, QStyle, QSystemTrayIcon, QMainWindow)

# submodules - export
from . import Isogeo2xlsx, Isogeo2docx

# #############################################################################
# ########## Globals ###############
# ##################################
current_locale = QLocale()
logger = logging.getLogger("isogeo2office")

# #############################################################################
# ######## QThreads ################
# ##################################

class AppPropertiesThread(QThread):
    signal = pyqtSignal(str)

    def __init__(self, api_manager: object):
        QThread.__init__(self)
        self.api_mngr = api_manager

    # run method gets called when we start the thread
    def run(self):
        """Get application and informations
        """
        # get application properties
        shares = self.api_mngr.isogeo.shares(token=self.api_mngr.token)
        # insert text
        text = "<html>"  # opening html content
        # Isogeo application authenticated in the plugin
        app = shares[0].get("applications")[0]
        text += "<p>{}<a href='{}' style='color: CornflowerBlue;'>{}</a> and "\
                .format(self.tr("This application is authenticated as "),
                        app.get("url", "https://isogeo.gitbooks.io/app-isogeo2office/content/"),
                        app.get("name", "Isogeo to Office"))
        # shares feeding the application
        if len(shares) == 1:
            text += "{}</p></br>".format(self.tr(" powered by 1 share:"))
        else:
            text += self.tr(" powered by {0} shares:</p></br>"
                            .format(len(shares))
                            )
        # shares details
        for share in shares:
            # share variables
            creator_name = share.get("_creator").get("contact").get("name")
            creator_email = share.get("_creator").get("contact").get("email")
            creator_id = share.get("_creator").get("_tag")[6:]
            share_url = "https://app.isogeo.com/groups/{}/admin/shares/{}"\
                        .format(creator_id, share.get("_id"))
            # formatting text
            text += "<p><a href='{}' style='color: CornflowerBlue;'><b>{}</b></a></p>"\
                    .format(share_url,
                            share.get("name"))
            text += "<p>{} {}</p>"\
                    .format(self.tr("Updated:"),
                            QDate.fromString(share.get("_modified")[:10],
                                             "yyyy-MM-dd").toString())

            text += "<p>{} {} - {}</p>"\
                    .format(self.tr("Contact:"),
                            creator_name,
                            creator_email)
            text += "<p><hr></p>"
        text += "</html>"
        # application and shares informations retrieved.
        # Now inform the main thread with the output (fill_app_props)
        self.signal.emit(text)



# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == "__main__":
    logging.critical("Can't be used as main script.")
