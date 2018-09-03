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
from datetime import datetime
import logging
from functools import partial
from logging.handlers import RotatingFileHandler
from os import listdir, path
import platform

# 3rd party library
from isogeo_pysdk import __version__ as pysdk_version
from PyQt5.QtCore import (QBasicTimer, QDate, QLocale, QSettings, QTimerEvent,
                          QTranslator, QThread, pyqtSignal, pyqtSlot)
from PyQt5.QtGui import QCloseEvent, QIcon
from PyQt5.QtWidgets import (QApplication, QComboBox, QDialog,
                             QMessageBox, QStyle, QSystemTrayIcon, QMainWindow)

import qdarkstyle
# submodules - export
from modules import isogeo2office_utils

# submodules - UI
from modules.ui.auth.auth_dlg import Auth
from modules.ui.credits.credits_dlg import Credits
from modules.ui.main.ui_win_IsogeoToOffice import Ui_win_IsogeoToOffice

# submodules - functional
from modules import IsogeoApiMngr
from modules import AppPropertiesThread, ExportExcelThread, ExportWordThread, ExportXmlThread

# #############################################################################
# ########## Globals ###############
# ##################################
app_dir = path.realpath(path.dirname(__file__))
app_logdir = path.join(app_dir, "_logs")
app_tpldir = path.join(app_dir, "templates")
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
class IsogeoToOffice_Main(QMainWindow):

    # attributes and global actions
    logger.info('OS: {0}'.format(platform.platform()))
    logger.info('Version: {0}'.format(__version__))
    logger.info('Isogeo PySDK version: {0}'.format(pysdk_version))
    logger.info('System locale: {0}'.format(current_locale.name()))

    # submodules shortcuts
    app_utils = isogeo2office_utils()

    def __init__(self):
        super().__init__()
        #self.ui = Ui_tabs_IsogeoToOffice()
        self.ui = Ui_win_IsogeoToOffice()
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

        self.ui.btn_launch_export.pressed.connect(partial(self.export))

        # -- Export tab - Output formats --------------------------------------
        self.ui.chb_output_excel.toggled\
               .connect(lambda: self.app_settings
                                    .setValue("formats/excel",
                                              int(self.ui.chb_output_excel.isChecked()
                                                  )
                                              )
                        )
        self.ui.chb_output_word.toggled\
               .connect(lambda: self.app_settings
                                    .setValue("formats/word",
                                              int(self.ui.chb_output_word.isChecked()
                                                  )
                                              )
                        )
        self.ui.chb_output_xml.toggled\
               .connect(lambda: self.app_settings
                                    .setValue("formats/xml",
                                              int(self.ui.chb_output_xml.isChecked()
                                                  )
                                              )
                        )
        # -- Settings tab - Export -------------------------------------------
        self.ui.btn_directory_change.pressed.connect(partial(self.set_output_folder))

        # populate Word templates combobox
        for tpl in listdir(app_tpldir):
            if path.splitext(tpl)[1].lower() == ".docx":
                self.ui.cbb_word_tpl.addItem(path.basename(tpl),
                                             path.join(app_tpldir, tpl))

        # -- Settings tab - Application authentication ------------------------
        # Change user -> see below for authentication form
        self.ui.btn_change_user.pressed.connect(partial(api_mngr.display_auth_form))
        # share text window
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

        # system tray icon
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon("img/favicon.ico"))

        # -- DISPLAY  ---------------------------------------------------------
        # shortcuts
        self.cbbs_filters = self.ui.grp_filters.findChildren(QComboBox)
        self.setWindowTitle("Isogeo to Office - v{}".format(__version__))
        self.show()
        self.init_api_connection()

    def displayer(self, ui_class):
        """A simple relay in charge of displaying independant UI classes."""
        ui_class.exec_()

    def init_api_connection(self):
        """After UI display, start to try to connect to Isogeo API.
        """
        self.processing(step="start")
        # check credentials
        if not api_mngr.manage_api_initialization():
            logger.error("No credentials")
            QMessageBox.warning(self,
                                self.tr("Authentication - Credentials missing"),
                                self.tr("Authentication to Isogeo API has failed."
                                        " Credentials seem to be missing.")
                                )
            return False
        else:
            logger.debug("Access granted. Fill the shares window")
            thread_app_props = AppPropertiesThread(api_mngr)
            thread_app_props.signal.connect(self.fill_app_props)
            thread_app_props.start()

        # launch empty search
        self.search(reset=1)

        # stop timer and progress bar
        self.processing(step='end')

    # -- SEARCH ---------------------------------------------------------------
    def search(self, reset: bool = 0):
        """Get filters and make search
        """
        self.processing("start")

        # fill
        if reset:
            logger.debug("Reset search form.")
            search = api_mngr.isogeo.search(api_mngr.token,
                                            page_size=0,
                                            whole_share=0,
                                            augment=1,
                                            tags_as_dicts=1)
        else:
            share_id, search_terms = self.get_selected_filters()
            logger.debug("Search with filters: {}. Share: {}."
                         .format(search_terms, share_id))
            search = api_mngr.isogeo.search(api_mngr.token,
                                            query=search_terms,
                                            share=share_id,
                                            page_size=0,
                                            whole_share=0,
                                            augment=1,
                                            tags_as_dicts=1)

        # update ui
        self.update_search_form(search)
        self.processing("end")

    def get_selected_filters(self):
        """Retrieve selected filters from the search form.
        """
        share_id = ""
        search_terms = ""
        for cbb in self.cbbs_filters:
            if cbb.itemData(cbb.currentIndex()).startswith("share:"):
                share_id = cbb.itemData(cbb.currentIndex()).split(":")[1]
            else:
                search_terms += cbb.itemData(cbb.currentIndex())

        return share_id, search_terms

    def update_search_form(self, search: dict):
        """Update search form with tags.
        """
        # query parameters
        logger.debug(search.get("query"))
        # COMBOBOXES - FILTERS
        # clear previous state
        for cbb in self.cbbs_filters:
            cbb.clear()
        tags = search.get("tags")
        # add none selection item
        for cbb in self.cbbs_filters:
            cbb.addItem(" - ", "")

        # Shares
        logger.debug(tags.keys())
        for k, v in tags.get("shares").items():
            self.ui.cbb_share.addItem(k, v)
        # Owners
        for k, v in tags.get("owners").items():
            self.ui.cbb_owner.addItem(k, v)
        # Types
        for k, v in tags.get("types").items():
            self.ui.cbb_type.addItem(k, v)
        # Keywords
        for k, v in tags.get("keywords").items():
            self.ui.cbb_keyword.addItem(k, v)

        # export button
        self.ui.btn_launch_export.setText(self.tr("Export {} metadata".format(search.get("total"))))

    # -- EXPORT ---------------------------------------------------------------
    def export(self):
        """Launch export"""
        # check export options
        li_opts = [self.ui.chb_output_excel.isChecked(),
                   self.ui.chb_output_word.isChecked(),
                   self.ui.chb_output_xml.isChecked()
                   ]
        if not any(li_opts):
            QMessageBox.critical(self,
                                 self.tr("Export option is missing"),
                                 self.tr("At least one export option required."))
            logger.error("No export option selected.")
            return
        else:
            pass

        # retrieve metadata to export
        self.processing("start")
        share_id, search_terms = self.get_selected_filters()
        includes = ["conditions",
                    "contacts",
                    "coordinate-system",
                    "events",
                    "feature-attributes",
                    "keywords",
                    "layers",
                    "limitations",
                    "links",
                    "operations",
                    "serviceLayers",
                    "specifications"]
        search_to_be_exported = api_mngr.isogeo.search(api_mngr.token,
                                                       query=search_terms,
                                                       share=share_id,
                                                       page_size=100,
                                                       whole_share=1,
                                                       include=includes,
                                                       check=0)
        self.processing("end")

        # prepare progress bar
        progbar_max = search_to_be_exported.get("total")
        self.ui.pgb_exports.setRange(0, progbar_max)
        self.ui.pgb_exports.reset()

        # -- File naming
        # prepare filepath
        generic_filepath = path.join(self.app_settings.value("settings/out_folder",
                                                             r"output/"),
                                     self.ui.txt_output_fileprefix.text()
                                     )
        # horodating ?
        opt_timestamp = self.ui.cbb_timestamp.currentText()
        logger.debug("Timestamp option: {}"
                     .format(opt_timestamp))
        if opt_timestamp == self.tr("No date (overwrite)"):
            horodatage = ""
        elif opt_timestamp == self.tr("Day"):
            dstamp = datetime.now()
            horodatage = "_{}-{}-{}".format(dstamp.year,
                                            dstamp.month,
                                            dstamp.day)
        elif opt_timestamp == self.tr("Datetime"):
            dstamp = datetime.now()
            horodatage = "_{}-{}-{}-{}{}{}".format(dstamp.year,
                                                   dstamp.month,
                                                   dstamp.day,
                                                   dstamp.hour,
                                                   dstamp.minute,
                                                   dstamp.second)
        else:
            logger.error("Timestamp option not recognized")
            horodatage = ""
        # metadata UUID
        opt_md_uuid = self.ui.int_md_uuid.value()
        logger.debug("UUID option: {}"
                     .format(opt_md_uuid))

        # EXCEL
        if self.ui.chb_output_excel.isChecked():
            logger.debug("Excel - Preparation")
            output_xlsx_filepath = "{}{}.xlsx".format(generic_filepath, horodatage)
            logger.debug("Excel - Destination file: {}".format(output_xlsx_filepath))
            self.thread_export_xlsx = ExportExcelThread(search_to_be_exported,
                                                        output_xlsx_filepath,
                                                        opt_attributes=self.ui.chb_xls_attributes.isChecked(),
                                                        opt_dasboard=self.ui.chb_xls_stats.isChecked())
            self.thread_export_xlsx.sig_step.connect(self.update_status_bar)
            self.thread_export_xlsx.start()
        else:
            pass

        # WORD
        if self.ui.chb_output_word.isChecked():
            logger.debug("Word - Preparation")
            output_docx_filepath = "{}{}".format(generic_filepath, horodatage)
            logger.debug("Word - Output folder: {}".format(output_docx_filepath))
            template_path = self.ui.cbb_word_tpl.itemData(self.ui.cbb_word_tpl.currentIndex())
            logger.debug("Word - Template choosen: {}".format(template_path))
            self.thread_export_docx = ExportWordThread(search_to_be_exported,
                                                       output_docx_filepath,
                                                       tpl_path=template_path,
                                                       timestamp=horodatage,
                                                       length_uuid=opt_md_uuid)
            self.thread_export_docx.sig_step.connect(self.update_status_bar)
            self.thread_export_docx.start()
        else:
            pass

        # XML
        if self.ui.chb_output_xml.isChecked():
            logger.debug("XML - Preparation")
            output_xml_filepath = "{}{}".format(generic_filepath, horodatage)
            logger.debug("XML - Output folder: {}".format(output_xml_filepath))
            self.thread_export_xml = ExportXmlThread(search_to_be_exported,
                                                     isogeo_api_mngr=api_mngr,
                                                     output_path=output_xml_filepath,
                                                     opt_zip=self.ui.chb_xml_zip.isChecked(),
                                                     timestamp=horodatage,
                                                     length_uuid=opt_md_uuid)
            self.thread_export_xml.sig_step.connect(self.update_status_bar)
            self.thread_export_xml.start()
        else:
            pass

    # -- UI utils -------------------------------------------------------------
    def closeEvent(self, event_sent):
        """Actions performed juste before UI is closed.

        :param QCloseEvent event_sent: event sent when the main UI is close
        """
        # misc
        self.app_settings.setValue("log/log_level", "10")

        # API
        self.app_settings.setValue("auth/app_id", api_mngr.api_app_id)
        self.app_settings.setValue("auth/app_secret", api_mngr.api_app_secret)
        self.app_settings.setValue("auth/url_base", api_mngr.api_url_base)
        self.app_settings.setValue("auth/url_auth", api_mngr.api_url_auth)
        self.app_settings.setValue("auth/url_token", api_mngr.api_url_token)
        self.app_settings.setValue(
            "auth/url_redirect", api_mngr.api_url_redirect)

        # output formats
        self.app_settings.setValue(
            "formats/excel", self.ui.chb_output_excel.isChecked())
        self.app_settings.setValue(
            "formats/word", self.ui.chb_output_word.isChecked())
        self.app_settings.setValue(
            "formats/xml", self.ui.chb_output_xml.isChecked())

        # location and naming rules
        # self.app_settings.setValue("settings/out_folder",
        #                            self.ui.lbl_output_folder_value.text())
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

    def processing(self, step: str = "start", max: int = 100):
        """Manage UI during a process: progress bar start/end, disable/enable
        widgets...

        :param str step: step of processing (start, end or progress)
        :param int max: length of the process
        """
        if step == "start":
            logger.debug("Start processing. Freezing search form.")
            self.ui.tab_export.setEnabled(False)
            self.timer.start(max, self)
        elif step == "end":
            logger.debug("End of process. Back to normal.")
            self.ui.tab_export.setEnabled(True)
            self.timer.stop()
        elif step == "progress":
            logger.debug("Progress")
        else:
            raise ValueError

    def set_output_folder(self):
        """Let user pick the folder where to store outputs.
        """
        # launch explorer
        selected_folder = self.app_utils.open_FileNameDialog(self,
                                                             file_type="folder",
                                                             from_dir=self.app_settings.value("settings/out_folder"))
        # test selected folder
        if not path.exists(selected_folder):
            logger.error("No folder selected")
            return False
        else:
            selected_folder = path.realpath(selected_folder)
            logger.debug("Output folder selected: {}".format(selected_folder))

        # fill label and setttings
        self.ui.lbl_output_folder_value.setText(path.basename(selected_folder))
        self.ui.lbl_output_folder_value.setToolTip(
            path.dirname(selected_folder))
        self.app_settings.setValue("settings/out_folder",
                                   selected_folder)

    def timerEvent(self, event_sent):
        """Timer event catcher in charge of updating the progress bar.

        :param QTimerEvent event_sent: event automatically sent by QBasicTimer
        """
        # check if step is over the end limit
        if self.step >= 100:
            self.timer.stop()
            return
        else:
            self.step += 1
            self.ui.pgb_exports.setValue(self.step)

    # -- UI Slots -------------------------------------------------------------
    @pyqtSlot(str)
    def fill_app_props(self, app_infos_retrieved: str = ""):
        """Get app properties and fillfull the share frame in settings tab.
        """
        self.ui.txt_shares.setText(app_infos_retrieved)
        # notification
        self.tray_icon.show()
        self.tray_icon.showMessage("Isogeo to Office",
                                   self.tr("Application information has been retrieved"),
                                   QIcon("img/favicon.ico"),
                                   2000
                                   )

    @pyqtSlot(int, str)
    def update_status_bar(self, prog_step: int = 1, status_msg: str = ""):
        """Display message into status bar
        """
        self.ui.lbl_statusbar.showMessage(status_msg)
        prog_val = self.ui.pgb_exports.value() + prog_step
        self.ui.pgb_exports.setValue(prog_val)


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
