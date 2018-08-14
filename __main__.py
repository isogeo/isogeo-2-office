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

# #############################################################################
# ########## Libraries #############
# ##################################

# standard library
from functools import partial
import logging
from logging.handlers import RotatingFileHandler
from os import path
import platform

# 3rd party library
from isogeo_pysdk import Isogeo, IsogeoChecker, __version__ as pysdk_version
from PyQt5 import QtWidgets
from PyQt5 import QtCore
import qdarkstyle

# submodules - UI
from modules.ui.auth.ui_authentication import Ui_dlg_authentication
from modules.ui.credits.ui_credits import Ui_dlg_credits
from modules.ui.main.ui_IsogeoToOffice import Ui_tabs_IsogeoToOffice

# submodules - functional
from modules import Isogeo2xlsx
from modules import Isogeo2docx
from modules import IsogeoStats
from modules import isogeo2office_utils

# #############################################################################
# ########## Globals ###############
# ##################################
app_dir = path.realpath(path.dirname(__file__))
app_logdir = path.join(app_dir, "_logs")
current_locale = QtCore.QLocale()

# VERSION
__version__ = "2.0.0-beta1"

# LOG FILE ##
logger = logging.getLogger("isogeo2office")
logging.captureWarnings(True)
logger.setLevel(logging.DEBUG)
log_form = logging.Formatter("%(asctime)s || %(levelname)s "
                             "|| %(module)s - %(lineno)d ||"
                             " %(funcName)s || %(message)s")
logfile = RotatingFileHandler(path.join(app_logdir,
                                        "log_IsogeoToOffice.log"),
                              "a", 5000000, 1)
logfile.setLevel(logging.DEBUG)
logfile.setFormatter(log_form)
logger.addHandler(logfile)
logger.info('================ Isogeo to office ===============')

# #############################################################################
# ########## Classes ###############
# ##################################
class IsogeoToOffice_Main(QtWidgets.QTabWidget):

    # attributes and global actions
    logger.info('OS: {0}'.format(platform.platform()))
    logger.info('Version: {0}'.format(__version__))
    logger.info('Isogeo PySDK version: {0}'.format(pysdk_version))
    logger.info('System locale: {0}'.format(current_locale.name()))

    # submodules shortcuts
    app_utils = isogeo2office_utils()

    def __init__(self):
        super().__init__()
        self.ui = Ui_tabs_IsogeoToOffice()
        self.ui.setupUi(self)
        self.initUI()

    def initUI(self):
        """Start UI display and widgets signals and slots.
        """
        # display
        self.show()

        """ --- CONNECTING UI WIDGETS <-> FUNCTIONS --- """
        self.ui.btn_reinit.pressed.connect(self.reinitialize_search)

        # -- Settings tab - Resources -----------------------------------------
        # report and log - see #53 and  #139
        self.ui.btn_log_dir.pressed.connect(partial(self.app_utils.open_dir_file, target=app_logdir))
        self.ui.btn_report.pressed.connect(
            partial(self.app_utils.open_urls,
                    li_url=["https://github.com/isogeo/isogeo-2-office/issues/new?title={}"
                            " - version {} Windows {}&labels=bug&milestone=3"
                            .format(self.tr("TITLE ISSUE REPORTED"),
                                    __version__,
                                    platform.platform()), ]
                    )
        )
        # help button
        self.ui.btn_help.pressed.connect(
            partial(self.app_utils.open_urls,
                    li_url=["https://isogeo.gitbooks.io/app-isogeo2office/", ]
                    )
        )
        # # view credits - see: #52
        # self.dockwidget.btn_credits.pressed.connect(partial(self.show_popup, popup='credits'))


    def init_api_connection(self):
        """After UI display, start to try to connect to Isogeo API.
        """
        # check credentials


        self.pgb_exports.start()


    def reinitialize_search(self):
        logger.debug("Search reset")
        self.ui.Export.setEnabled(False)
        logger.debug("poupoupidou, resetting form")
        self.ui.Export.setEnabled(True)


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
