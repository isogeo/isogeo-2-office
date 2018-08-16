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

import logging
import platform
# standard library
from functools import partial
from logging.handlers import RotatingFileHandler
from os import path

# 3rd party library
from isogeo_pysdk import Isogeo, IsogeoChecker
from isogeo_pysdk import __version__ as pysdk_version
from PyQt5.QtCore import QLocale, QSettings, QBasicTimer, QTranslator
from PyQt5.QtWidgets import (QApplication, QDialog, QMenu, QStyle, QSystemTrayIcon,
                             QTabWidget)
import qdarkstyle

# submodules - functional
from modules.utils.api import IsogeoApiMngr
from modules import Isogeo2docx, Isogeo2xlsx, IsogeoStats, isogeo2office_utils
# submodules - UI
from modules.ui.auth.auth_dlg import Auth
from modules.ui.credits.credits_dlg import Credits
from modules.ui.main.ui_IsogeoToOffice import Ui_tabs_IsogeoToOffice

# #############################################################################
# ########## Globals ###############
# ##################################
app_dir = path.realpath(path.dirname(__file__))
app_logdir = path.join(app_dir, "_logs")
current_locale = QLocale()

api_mngr = IsogeoApiMngr()

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
class IsogeoToOffice_Main(QTabWidget):

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
        # Settings
        self.app_settings = QSettings('Isogeo', 'IsogeoToOffice')
        # usage metrics
        launch_counter = self.app_settings.value("usage/launch", 0)
        self.app_settings.setValue("usage/launch", launch_counter + 1)
        # Credits
        self.ui_credits = Credits()
        # Auth
        api_mngr.ui_auth_form = Auth()
        api_mngr.auth_folder = path.join(app_dir, "_auth")
        # build UI
        self.initUI()

    def initUI(self):
        """Start UI display and widgets signals and slots.
        """
        # timer and progress bar
        self.timer = QBasicTimer()
        self.step = 0

        """ --- CONNECTING UI WIDGETS <-> FUNCTIONS --- """
        # -- Export tab - Filters ---------------------------------------------
        self.ui.cbb_share.activated.connect(partial(self.search, 0))
        self.ui.cbb_type.activated.connect(partial(self.search, 0))
        self.ui.cbb_owner.activated.connect(partial(self.search, 0))
        self.ui.cbb_keyword.activated.connect(partial(self.search, 0))
        self.ui.btn_reinit.pressed.connect(partial(self.search, 1))

        # -- Settings tab - Application authentication ------------------------
        # Change user -> see below for authentication form
        self.ui.btn_change_user.pressed.connect(partial(api_mngr.display_auth_form))
        # share text window
        self.ui.txt_shares.setOpenLinks(False)
        self.ui.txt_shares.anchorClicked.connect(self.app_utils.open_urls)

        # -- Settings tab - Resources -----------------------------------------
        self.ui.btn_log_dir.pressed.connect(partial(self.app_utils.open_dir_file,
                                                    target=app_logdir))
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
        # view credits
        self.ui.btn_credits.pressed.connect(partial(self.displayer,
                                                    self.ui_credits))

        # -- DISPLAY then check API
        self.show()
        self.init_api_connection()

    def displayer(self, ui_class):
        """A simple relay in charge of displaying independant UI classes."""
        ui_class.exec_()

    def init_api_connection(self):
        """After UI display, start to try to connect to Isogeo API.
        """
        # check credentials
        self.processing("start")
        if not api_mngr.manage_api_initialization():
            logger.error("No credentials")
        
        # 
        #self.isogeo = Isogeo()

        #self.processing("stop")
        #


    def search(self, reset: bool = 0):
        """Get filters and make search
        """
        if reset:
            logger.debug("Reset search form.")

        else:
            logger.debug("Search with filters")
            share= self.ui.cbb_share.currentText()
            type = self.ui.cbb_type.currentText()
            keyword = self.ui.cbb_keyword.currentText()
            owner = self.ui.cbb_owner.currentText()
        
            print(share, type, keyword, owner)


    def export(self):
        """Launch export"""
        print(self.app_settings.allKeys())

    # -- UI utils -------------------------------------------------------------
    def processing(self, step: str = "start"):
        """Manage UI during a process: progress bar start/end, disable/enable
        widgets...

        :param str step: step of processing (start, end or progress)
        """
        if step == "start":
            logger.debug("Start processing. Freezing search form.")
            self.ui.Export.setEnabled(False)
            self.timer.start(100, self)
        elif step == "end":
            logger.debug("End of process. Back to normal.")
            self.ui.Export.setEnabled(True)
            self.timer.stop()
        elif step == "progress":
            logger.debug("Progress")
        else:
            raise ValueError

    def closeEvent(self, event_sent):
        """Actions performed juste before UI is closed."""
        # misc
        self.app_settings.setValue("log/log_level", "10")

        # API
        self.app_settings.setValue("auth/app_id", api_mngr.api_app_id)
        self.app_settings.setValue("auth/app_secret", api_mngr.api_app_secret)
        self.app_settings.setValue("auth/url_base", api_mngr.api_url_base)
        self.app_settings.setValue("auth/url_auth", api_mngr.api_url_auth)
        self.app_settings.setValue("auth/url_token", api_mngr.api_url_token)
        self.app_settings.setValue("auth/url_redirect", api_mngr.api_url_redirect)

        # output formats
        self.app_settings.setValue("formats/excel", self.ui.chb_output_excel.isChecked())
        self.app_settings.setValue("formats/word", self.ui.chb_output_word.isChecked())
        self.app_settings.setValue("formats/xml", self.ui.chb_output_xml.isChecked())

        # location and naming rules
        self.app_settings.setValue("settings/out_folder",
                                   self.ui.lbl_output_folder_value.text())
        self.app_settings.setValue("settings/out_prefix",
                                   self.ui.txt_output_fileprefix.text())
        self.app_settings.setValue("settings/timestamps",
                                   self.ui.cbb_timestamp.currentText())
        self.app_settings.setValue("settings/uuid_length",
                                   self.ui.int_md_uuid.text())
        
        # export options
        self.app_settings.setValue("settings/xls_sheet_attributes",
                                   self.ui.chb_xls_attributes.isChecked())
        self.app_settings.setValue("settings/xls_sheet_dashboard",
                                   self.ui.chb_xls_stats.isChecked())
        self.app_settings.setValue("settings/doc_tpl_name",
                                   self.ui.cbb_word_tpl.currentText())
        self.app_settings.setValue("settings/xml_zip",
                                   self.ui.chb_xml_zip.isChecked())

        # accept the close
        event_sent.accept()

    def timerEvent(self, event_sent):

        if self.step >= 100:
            self.timer.stop()
            return

        self.step += 1
        self.ui.pgb_exports.setValue(self.step)

# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == "__main__":
    import sys
    # create the application and the main window
    app = QApplication(sys.argv)
    # apply dark style
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    # apply language
    locale_path = path.join(app_dir,
                            'i18n',
                            'IsogeoToOffice_{}.qm'.format(current_locale.system().name()))
    translator = QTranslator()
    translator.load(path.realpath(locale_path))
    app.installTranslator(translator)
    # link to Isogeo to Office main UI
    i2o = IsogeoToOffice_Main()
    sys.exit(app.exec_())