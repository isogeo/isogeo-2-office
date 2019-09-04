# -*- coding: UTF-8 -*-
#! python3

"""
    Isogeo To Office - Main launcher

    Purpose:     Get metadatas from Isogeo and export to desktop formats
    Author:      Isogeo
    Python:      3.7.x
"""

# #############################################################################
# ########## Libraries #############
# ##################################

# standard library
import logging
import pathlib  # TO DO: replace os.path by pathlib
import platform
from functools import partial
from logging.handlers import RotatingFileHandler
from os import listdir, path

# 3rd party
import qdarkstyle
import semver
from dotenv import load_dotenv
from isogeo_pysdk import __version__ as pysdk_version, MetadataSearch

# PyQt
from PyQt5.QtCore import QLocale, QSettings, QThread, QTranslator, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QCloseEvent, QIcon
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QMainWindow,
    QMessageBox,
    QStyleFactory,
)

# submodules - functional
from modules import (
    IsogeoApiMngr,
    ThreadAppProperties,
    ThreadExportExcel,
    ThreadExportWord,
    ThreadExportXml,
    ThreadSearch,
    ThreadThumbnails,
    isogeo2office_utils,
)

# submodules - UI
from modules.ui.auth.auth_dlg import Auth
from modules.ui.credits.credits_dlg import Credits
from modules.ui.main.ui_win_IsogeoToOffice import Ui_win_IsogeoToOffice
from modules.ui.systray.ui_systraymenu import SystrayMenu

# #############################################################################
# ########## Globals ###############
# ##################################

# load specific enviroment vars
load_dotenv(".env")

# required subfolders
pathlib.Path("_auth/").mkdir(exist_ok=True)
pathlib.Path("_logs/").mkdir(exist_ok=True)
pathlib.Path("_input/").mkdir(exist_ok=True)
pathlib.Path("_output/").mkdir(exist_ok=True)
pathlib.Path("_templates/").mkdir(exist_ok=True)
pathlib.Path("_thumbnails/").mkdir(exist_ok=True)

# vars
app_dir = path.realpath(path.dirname(__file__))
app_logdir = path.join(app_dir, "_logs")
app_outdir = path.join(app_dir, "_output")
app_thbdir = path.join(app_dir, "_thumbnails")
app_tpldir = path.join(app_dir, "_templates")
current_locale = QLocale()

api_mngr = IsogeoApiMngr()

# VERSION
__version__ = "2.1.0-beta1"

# LOG FILE #
# log level depends on version
if "beta" in __version__:
    log_level = logging.DEBUG
else:
    log_level = logging.INFO

logger = logging.getLogger("isogeo2office")
logging.captureWarnings(True)
logger.setLevel(log_level)
log_form = logging.Formatter(
    "%(asctime)s || %(levelname)s "
    "|| %(module)s - %(lineno)d ||"
    " %(funcName)s || %(message)s"
)
logfile = RotatingFileHandler(
    path.join(app_logdir, "log_IsogeoToOffice.log"), "a", 5000000, 1
)
logfile.setLevel(log_level)
logfile.setFormatter(log_form)

# info to the console
log_console_handler = logging.StreamHandler()
log_console_handler.setLevel(logging.INFO)
log_console_handler.setFormatter(log_form)

logger.addHandler(log_console_handler)
logger.addHandler(logfile)
logger.info("================ Isogeo to office ===============")


# #############################################################################
# ########## Classes ###############
# ##################################
class IsogeoToOffice_Main(QMainWindow):

    # attributes and global actions
    logger.info("OS: {0}".format(platform.platform()))
    logger.info("Version: {0}".format(__version__))
    logger.info("Isogeo PySDK version: {0}".format(pysdk_version))
    logger.info("System locale: {0}".format(current_locale.name()))

    # submodules shortcuts
    app_utils = isogeo2office_utils()

    def __init__(self):
        super().__init__()
        self.ui = Ui_win_IsogeoToOffice()
        self.ui.setupUi(self)
        # Settings
        self.app_settings = QSettings("Isogeo", "IsogeoToOffice")
        self.settings_noSave = 0
        # usage
        launch_counter = self.app_settings.value("usage/launch", 0)
        self.app_settings.setValue("usage/launch", launch_counter + 1)
        self.app_settings.setValue("usage/version", __version__)
        # Credits
        self.ui_credits = Credits()
        # Auth
        api_mngr.auth_form_request_url = self.tr(
            "https://pipedrivewebforms.com/form/b5bdbdb9b34c3c61202cd8414accbbe252944"
        )
        api_mngr.ui_auth_form = Auth()
        api_mngr.auth_folder = path.join(app_dir, "_auth")
        self.app_utils.clean_credentials_files(path.join(app_dir, "_auth"))
        # build UI
        self.initUI()

    def initUI(self):
        """Start UI display and widgets signals and slots.
        """
        """ --- CONNECTING UI WIDGETS <-> FUNCTIONS --- """
        # -- Export tab - Filters ---------------------------------------------
        self.ui.cbb_share.activated.connect(partial(self.search, "update"))
        self.ui.cbb_type.activated.connect(partial(self.search, "update"))
        self.ui.cbb_owner.activated.connect(partial(self.search, "update"))
        self.ui.cbb_keyword.activated.connect(partial(self.search, "update"))
        self.ui.btn_reinit.pressed.connect(partial(self.search, "reset"))

        self.ui.btn_launch_export.pressed.connect(partial(self.search, "export"))

        # -- Export tab - Output formats --------------------------------------
        self.ui.chb_output_excel.toggled.connect(
            lambda: self.app_settings.setValue(
                "formats/excel", int(self.ui.chb_output_excel.isChecked())
            )
        )
        self.ui.chb_output_word.toggled.connect(
            lambda: self.app_settings.setValue(
                "formats/word", int(self.ui.chb_output_word.isChecked())
            )
        )
        self.ui.chb_output_xml.toggled.connect(
            lambda: self.app_settings.setValue(
                "formats/xml", int(self.ui.chb_output_xml.isChecked())
            )
        )
        # -- Settings tab - Export -------------------------------------------
        self.ui.btn_directory_change.pressed.connect(partial(self.set_output_folder))
        self.ui.btn_thumbnails_edit.pressed.connect(
            partial(
                self.app_utils.open_dir_file,
                path.join(app_dir, "_thumbnails/thumbnails.xlsx"),
            )
        )

        # -- Settings tab - Timestamp -----------------------------------------
        self.ui.rdb_timestamp_no.toggled.connect(
            lambda: self.app_settings.setValue("settings/timestamps", "no")
        )
        self.ui.rdb_timestamp_day.toggled.connect(
            lambda: self.app_settings.setValue("settings/timestamps", "day")
        )
        self.ui.rdb_timestamp_datetime.toggled.connect(
            lambda: self.app_settings.setValue("settings/timestamps", "datetime")
        )
        # -- Settings tab - Word ----------------------------------------------
        for tpl in listdir(app_tpldir):
            if path.splitext(tpl)[1].lower() == ".docx":
                self.ui.cbb_word_tpl.addItem(
                    path.basename(tpl), path.join(app_tpldir, tpl)
                )

        self.ui.btn_thumbnails_update.pressed.connect(
            partial(self.search, "thumbnails")
        )
        # -- Settings tab - System tray icon ----------------------------------
        self.ui.chb_systray_minimize.toggled.connect(
            lambda: self.app_settings.setValue(
                "settings/systray_minimize",
                int(self.ui.chb_systray_minimize.isChecked()),
            )
        )

        # -- Settings tab - Application authentication ------------------------
        # Change user -> see below for authentication form
        self.ui.btn_change_user.pressed.connect(partial(api_mngr.display_auth_form))
        api_mngr.ui_auth_form.btn_browse_credentials.pressed.connect(
            partial(api_mngr.credentials_uploader)
        )
        api_mngr.ui_auth_form.btn_ok_cancel.pressed.connect(self.update_credentials)
        # share text window
        self.ui.txt_shares.anchorClicked.connect(self.app_utils.open_urls)

        # -- Settings tab - Resources -----------------------------------------
        self.ui.btn_log_dir.pressed.connect(
            partial(self.app_utils.open_dir_file, target=app_logdir)
        )
        self.ui.btn_report.pressed.connect(
            partial(
                self.app_utils.open_urls,
                li_url=[
                    "https://github.com/isogeo/isogeo-2-office/issues/new?title={}"
                    " - version {} Windows {}&labels=bug&milestone=3".format(
                        self.tr("TITLE ISSUE REPORTED"),
                        __version__,
                        platform.platform(),
                    )
                ],
            )
        )
        # help button
        self.ui.btn_help.pressed.connect(
            partial(
                self.app_utils.open_urls,
                li_url=["http://help.isogeo.com/isogeo2office/"],
            )
        )
        # reset factory defaults
        self.ui.btn_settings_reset.pressed.connect(partial(self.settings_reset))

        # view credits
        self.ui.btn_credits.pressed.connect(partial(self.displayer, self.ui_credits))

        # system tray icon
        self.tray_icon = SystrayMenu(parent=self)
        self.tray_icon.setIcon(QIcon("resources/icon.png"))
        self.tray_icon.act_quit.triggered.connect(self.close)
        self.tray_icon.act_show.triggered.connect(self.show)
        self.tray_icon.act_hide.triggered.connect(self.hide)

        # -- DISPLAY  ---------------------------------------------------------
        # shortcuts
        self.cbbs_filters = self.ui.grp_filters.findChildren(QComboBox)
        self.setWindowTitle("Isogeo to Office - v{}".format(__version__))
        try:
            self.settings_loader()
        except TypeError as e:
            self.app_settings.remove("settings")
            logger.error(
                "Settings loading failed: {}. Settings have been reset.".format(e)
            )
        self.show()
        self.init_api_connection()

    def init_api_connection(self):
        """After UI display, start to try to connect to Isogeo API.
        """
        self.processing(step="start")
        # check credentials
        if not api_mngr.manage_api_initialization():
            logger.error("Connection to Isogeo API failed.")
            QMessageBox.warning(
                self,
                self.tr("Authentication - Credentials missing"),
                self.tr(
                    "Authentication to Isogeo API has failed."
                    " Credentials seem to be missing."
                ),
            )
            return False
        else:
            logger.debug("Access granted. Fill the shares window")
            self.thread_app_props = ThreadAppProperties(api_mngr)
            self.thread_app_props.sig_finished.connect(self.fill_app_props)
            self.thread_app_props.start()

        # instanciate search thread
        self.thread_search = ThreadSearch(api_mngr)
        # launch empty search
        self.search(search_type="reset")

    def settings_loader(self):
        """Load application settings from QSettings and update UI with."""
        # output formats
        self.ui.chb_output_excel.setChecked(
            self.app_settings.value("formats/excel", False, type=bool)
        )
        self.ui.chb_output_word.setChecked(
            self.app_settings.value("formats/word", False, type=bool)
        )
        self.ui.chb_output_xml.setChecked(
            self.app_settings.value("formats/xml", False, type=bool)
        )
        # location and naming rules
        self.ui.lbl_output_folder_value.setText(
            self.app_settings.value("settings/out_folder_label", "_output")
        )
        path_output_folder = self.app_settings.value(
            "settings/out_folder_path", app_outdir
        )
        self.ui.lbl_output_folder_value.setToolTip(path_output_folder)
        self.ui.btn_open_output_folder.pressed.connect(
            partial(self.app_utils.open_dir_file, path_output_folder)
        )

        self.ui.txt_output_fileprefix.setText(
            self.app_settings.value("settings/out_prefix", "Isogeo_")
        )
        self.ui.int_md_uuid.setValue(
            self.app_settings.value("settings/uuid_length", 5, type=int)
        )
        # export options
        self.ui.chb_xls_attributes.setChecked(
            self.app_settings.value("settings/xls_sheet_attributes", True, type=bool)
        )
        self.ui.chb_xls_stats.setChecked(
            self.app_settings.value("settings/xls_sheet_dashboard", True, type=bool)
        )
        self.ui.chb_xml_zip.setChecked(
            self.app_settings.value("settings/xml_zip", False, type=bool)
        )
        tpl_index = self.ui.cbb_word_tpl.findText(
            self.app_settings.value("settings/doc_tpl_name", "template_Isogeo.docx")
        )
        self.ui.cbb_word_tpl.setCurrentIndex(tpl_index)

        # timestamps
        if self.app_settings.value("settings/timestamps", type=str) == "no":
            self.ui.rdb_timestamp_no.setChecked(1)
        elif self.app_settings.value("settings/timestamps", type=str) == "day":
            self.ui.rdb_timestamp_day.setChecked(1)
        elif self.app_settings.value("settings/timestamps", type=str) == "datetime":
            self.ui.rdb_timestamp_datetime.setChecked(1)
        else:
            logger.warning(
                "Timestamp option not recognized: {}".format(
                    self.app_settings.value("settings/timestamps")
                )
            )

        # misc
        self.ui.chb_systray_minimize.setChecked(
            self.app_settings.value("settings/systray_minimize", False, type=bool)
        )

        # try full restore
        try:
            self.restoreGeometry(self.app_settings.value("settings/geometry"))
            self.restoreState(self.app_settings.value("settings/windowState"))
            logger.debug("Application restore successed.")
        except AttributeError:
            logger.debug("Application restore failed.")
        # end of method
        logger.debug("Settings loaded")

    # -- SEARCH ---------------------------------------------------------------
    def search(self, search_type: str = "update"):
        """Get filters and make search.

        :param str search_type: can be update, reset or export. Defaults to 'update'.
        """
        # configure thread search
        self.ui.pgb_exports.setRange(0, 0)
        self.thread_search.search_type = search_type

        # depending on search type
        if search_type == "reset":
            self.thread_search.search_params = {
                "page_size": 0,
                "whole_results": 0,
                "augment": 1,
                "tags_as_dicts": 1,
            }
            self.thread_search.sig_finished.connect(self.update_search_form)
            logger.info("Search  prepared - {}".format(search_type.upper()))
        elif search_type == "update":
            share_id, search_terms = self.get_selected_filters()
            self.thread_search.search_params = {
                "query": search_terms,
                "share": share_id,
                "page_size": 0,
                "whole_results": 0,
                "augment": 1,
                "tags_as_dicts": 1,
            }
            self.thread_search.sig_finished.disconnect()
            self.thread_search.sig_finished.connect(self.update_search_form)
            logger.info("Search  prepared - {}".format(search_type.upper()))
        elif search_type == "export":
            # checks
            if not self.export_check():
                self.processing("end", 100)
                return False

            # prepare search and launch export process
            share_id, search_terms = self.get_selected_filters()
            self.thread_search.search_params = {
                "query": search_terms,
                "share": share_id,
                "whole_results": 1,
                "include": "all",
                "check": 0,
            }
            self.thread_search.sig_finished.disconnect()
            self.thread_search.sig_finished.connect(self.export_process)
            logger.info("Search  prepared - {}".format(search_type.upper()))
        elif search_type == "thumbnails":
            # checks
            if not self.export_check():
                self.processing("end", 100)
                return False

            # prepare search and launch export process
            share_id, search_terms = self.get_selected_filters()
            self.thread_search.search_params = {
                "query": search_terms,
                "share": share_id,
                "page_size": 100,
                "whole_results": 1,
                "check": 0,
            }
            self.thread_search.sig_finished.disconnect()
            self.thread_search.sig_finished.connect(self.thumbnails_generation)
            logger.info("Search  prepared - {}".format(search_type.upper()))
        else:
            raise ValueError

        # finally, start thread
        self.search_type = search_type
        self.thread_search.start()
        self.update_status_bar(
            prog_step=0, status_msg=self.tr("Waiting for Isogeo API")
        )

    def get_selected_filters(self) -> tuple:
        """Retrieve selected filters from the search form.

        :rtype: tuple
        """
        share_id = ""
        search_terms = ""
        for cbb in self.cbbs_filters:
            if cbb.itemData(cbb.currentIndex()).startswith("share:"):
                share_id = cbb.itemData(cbb.currentIndex()).split(":")[1]
            else:
                search_terms += cbb.itemData(cbb.currentIndex()) + " "

        return share_id, search_terms.strip()

    # -- EXPORT ---------------------------------------------------------------
    def export_check(self) -> bool:
        """Performs checks before export."""
        # check export options
        self.li_opts = [
            self.ui.chb_output_excel.isChecked(),
            self.ui.chb_output_word.isChecked(),
            self.ui.chb_output_xml.isChecked(),
        ]
        if not any(self.li_opts):
            QMessageBox.critical(
                self,
                self.tr("Export option is missing"),
                self.tr("At least one export option required."),
            )
            logger.error("No export option selected.")
            return False
        else:
            logger.debug(
                "Export check - {} output formats selected".format(sum(self.li_opts))
            )
            return True

    @pyqtSlot(MetadataSearch)
    def export_process(self, search_to_be_exported: MetadataSearch):
        """Export each metadata in checked output formats.

        :param MetadataSearch search_to_be_exported: Isogeo search response to export
        """
        # minimize application during process if asked. See #22
        if self.ui.chb_systray_minimize.isChecked():
            self.tray_icon.act_hide.trigger()
        # prepare progress bar
        progbar_max = sum(self.li_opts) * search_to_be_exported.total
        self.ui.pgb_exports.setRange(1, progbar_max)
        self.ui.pgb_exports.reset()

        # -- File naming
        # prepare filepath
        generic_filepath = path.join(
            self.app_settings.value("settings/out_folder_path", app_outdir),
            self.ui.txt_output_fileprefix.text(),
        )
        # horodating ?
        opt_timestamp = self.app_settings.value("settings/timestamps", "no")
        logger.debug("Timestamp option: {}".format(opt_timestamp))
        horodatage = self.app_utils.timestamps_picker(opt_timestamp)
        logger.debug("Timestamp value applied: {}".format(horodatage))
        # metadata UUID
        opt_md_uuid = self.ui.int_md_uuid.value()
        logger.debug("UUID option: {}".format(opt_md_uuid))

        # -- Inputs ---
        thumbs_filepath = path.join(app_thbdir, "thumbnails.xlsx")
        try:
            thumbnails_loaded = self.app_utils.thumbnails_mngr(thumbs_filepath)
        except Exception as e:
            logger.error(e)
            thumbnails_loaded = {None: (None, None)}

        # EXCEL
        if self.ui.chb_output_excel.isChecked():
            logger.debug("Excel - Preparation")
            output_xlsx_filepath = "{}{}.xlsx".format(generic_filepath, horodatage)
            logger.debug("Excel - Destination file: {}".format(output_xlsx_filepath))
            self.thread_export_xlsx = ThreadExportExcel(
                search_to_be_exported,
                output_xlsx_filepath,
                url_base_edit=api_mngr.isogeo.app_url,
                url_base_view=api_mngr.isogeo.oc_url,
                shares=api_mngr.isogeo._shares,
                opt_attributes=self.ui.chb_xls_attributes.isChecked(),
                opt_dasboard=self.ui.chb_xls_stats.isChecked(),
                opt_fillfull=0,
                opt_inspire=0,
            )
            self.thread_export_xlsx.sig_step.connect(self.update_status_bar)
            self.thread_export_xlsx.start()
        else:
            pass

        # WORD
        if self.ui.chb_output_word.isChecked():
            logger.debug("Word - Preparation")
            # output folder
            output_docx_filepath = "{}{}".format(generic_filepath, horodatage)
            logger.debug("Word - Output folder: {}".format(output_docx_filepath))
            # template
            template_path = self.ui.cbb_word_tpl.itemData(
                self.ui.cbb_word_tpl.currentIndex()
            )
            if not template_path or not path.exists(template_path):
                logger.warning(
                    "Word - No template choosen. trying to use the Isogeo default."
                )
                if not path.exists(path.join(app_tpldir, "template_Isogeo.docx")):
                    self.update_status_bar(0, self.tr("Word - Error: no template."))
                    logger.error("Word - No available template")
                    return False
                else:
                    template_path = path.join(app_tpldir, "template_Isogeo.docx")
            else:
                logger.debug("Word - Template choosen: {}".format(template_path))

            # instanciate thread
            self.thread_export_docx = ThreadExportWord(
                search_to_be_exported,
                output_docx_filepath,
                tpl_path=template_path,
                url_base_edit=api_mngr.isogeo.app_url,
                url_base_view=api_mngr.isogeo.oc_url,
                shares=api_mngr.isogeo._shares,
                thumbnails=thumbnails_loaded,
                timestamp=horodatage,
                length_uuid=opt_md_uuid,
            )
            self.thread_export_docx.sig_step.connect(self.update_status_bar)
            self.thread_export_docx.start()
        else:
            pass

        # XML
        if self.ui.chb_output_xml.isChecked():
            logger.debug("XML - Preparation")
            output_xml_filepath = "{}{}".format(generic_filepath, horodatage)
            logger.debug("XML - Output folder: {}".format(output_xml_filepath))
            self.thread_export_xml = ThreadExportXml(
                search_to_be_exported,
                isogeo_api_mngr=api_mngr,
                output_path=output_xml_filepath,
                opt_zip=self.ui.chb_xml_zip.isChecked(),
                timestamp=horodatage,
                length_uuid=opt_md_uuid,
            )
            self.thread_export_xml.sig_step.connect(self.update_status_bar)
            self.thread_export_xml.start()
        else:
            pass

    @pyqtSlot(MetadataSearch)
    def thumbnails_generation(self, search_to_be_exported: MetadataSearch):
        """Process thumbnails table generation.

        :param MetadataSearch search_to_be_exported: Isogeo search response to export
        """
        # prepare progress bar
        progbar_max = sum(self.li_opts) * search_to_be_exported.total
        self.ui.pgb_exports.setRange(1, progbar_max)
        self.ui.pgb_exports.reset()

        # Try to load from existig table
        thumbs_filepath = path.join(app_thbdir, "thumbnails.xlsx")
        try:
            thumbnails_loaded = self.app_utils.thumbnails_mngr(thumbs_filepath)
        except IOError as e:
            QMessageBox.critical(
                self,
                self.tr("Thumbnails - Table already opened"),
                self.tr(
                    "The thumbnails matching table {} is "
                    "already opened. Close it please "
                    "before to try again.".format(thumbs_filepath)
                ),
            )
            self.update_status_bar(
                0,
                self.tr(
                    "Error - Thumbnails table is opened. Close it before contiinue."
                ),
            )
            return False
        except KeyError as e:
            QMessageBox.warning(
                self,
                self.tr("Thumbnails - Table structure"),
                self.tr(
                    "The thumbnails matching table {} is "
                    "not compliant with the expected structure."
                    "{}{} {}".format(
                        thumbs_filepath,
                        self.tr(
                            "\nA new table will be created but "
                            "previous data will be lost."
                        ),
                        self.tr("\nError message:"),
                        e,
                    )
                ),
            )
            thumbnails_loaded = {None: None}
        except Exception as e:
            QMessageBox.warning(
                self,
                self.tr("Thumbnails - Unknown error"),
                self.tr(
                    "An unknown error occurred reading the"
                    "thumbnails matching table {}. "
                    "Please report it."
                    "{}{} {}".format(
                        thumbs_filepath,
                        self.tr(
                            "\nA new table will be created but "
                            "previous data will be lost."
                        ),
                        self.tr("\nError message:"),
                        e,
                    )
                ),
            )
            thumbnails_loaded = {None: (None, None)}

        # STORE
        logger.debug("Thumbnails - Preparation")
        logger.debug("Thumbnails - Output file: {}".format(thumbs_filepath))
        self.thread_thumbnails_gen = ThreadThumbnails(
            search_to_be_exported,
            output_path=thumbs_filepath,
            thumbnails=thumbnails_loaded,
        )
        self.thread_thumbnails_gen.sig_step.connect(self.update_status_bar)
        self.thread_thumbnails_gen.start()

    # -- UI utils -------------------------------------------------------------
    def closeEvent(self, event_sent):
        """Actions performed juste before UI is closed.

        :param QCloseEvent event_sent: event sent when the main UI is close
        """
        logger.debug("Closing application...")
        # force remove UI elements
        self.tray_icon.hide()
        self.tray_icon.deleteLater()

        if not self.settings_noSave:
            # -- Save settings ------------------------------------------------
            self.app_settings.setValue("settings/log_level", logger.level)
            # API
            self.app_settings.setValue("auth/app_id", api_mngr.api_app_id)
            self.app_settings.setValue("auth/app_secret", api_mngr.api_app_secret)
            self.app_settings.setValue("auth/app_type", api_mngr.api_app_type)
            self.app_settings.setValue("auth/platform", api_mngr.api_platform)
            self.app_settings.setValue("auth/url_base", api_mngr.api_url_base)
            self.app_settings.setValue("auth/url_auth", api_mngr.api_url_auth)
            self.app_settings.setValue("auth/url_token", api_mngr.api_url_token)
            self.app_settings.setValue("auth/url_redirect", api_mngr.api_url_redirect)

            # output formats
            self.app_settings.setValue(
                "formats/excel", self.ui.chb_output_excel.isChecked()
            )
            self.app_settings.setValue(
                "formats/word", self.ui.chb_output_word.isChecked()
            )
            self.app_settings.setValue(
                "formats/xml", self.ui.chb_output_xml.isChecked()
            )

            # location and naming rules
            # output folder is defined by its own method 'set_output_folder'
            self.app_settings.setValue(
                "settings/out_prefix", self.ui.txt_output_fileprefix.text()
            )
            self.app_settings.setValue(
                "settings/uuid_length", self.ui.int_md_uuid.value()
            )

            # export options
            self.app_settings.setValue(
                "settings/xls_sheet_attributes", self.ui.chb_xls_attributes.isChecked()
            )
            self.app_settings.setValue(
                "settings/xls_sheet_dashboard", self.ui.chb_xls_stats.isChecked()
            )
            self.app_settings.setValue(
                "settings/doc_tpl_name", self.ui.cbb_word_tpl.currentText()
            )
            self.app_settings.setValue(
                "settings/xml_zip", self.ui.chb_xml_zip.isChecked()
            )

            # misc
            self.app_settings.setValue(
                "settings/systray_minimize", self.ui.chb_systray_minimize.isChecked()
            )

            # global UI position
            self.app_settings.setValue("settings/geometry", self.saveGeometry())

            # global UI position
            self.app_settings.setValue("settings/geometry", self.saveGeometry())
            self.app_settings.setValue("settings/windowState", self.saveState())

        # accept the close
        logger.debug("See you!")
        event_sent.accept()

    def displayer(self, ui_class):
        """A simple relay in charge of displaying independant UI classes."""
        ui_class.exec_()

    def processing(self, step: str = "start", progbar_max: int = 0):
        """Manage UI during a process: progress bar start/end, disable/enable
        widgets.

        :param str step: step of processing (start, end or progress)
        :param int progbar_max: maximum range for the progress bar
        """
        if step == "start":
            logger.debug("Start processing. Freezing search form.")
            self.ui.tab_export.setEnabled(False)
        elif step == "end":
            logger.debug("End of process. Back to normal.")
            self.ui.pgb_exports.setRange(0, progbar_max)
            self.ui.tab_export.setEnabled(True)
            self.ui.btn_launch_export.setEnabled(True)
        elif step == "progress":
            logger.debug("Progress")
        else:
            raise ValueError

    def set_output_folder(self):
        """Let user pick the folder where to store outputs.
        """
        # launch explorer
        selected_folder = self.app_utils.open_FileNameDialog(
            self,
            file_type="folder",
            from_dir=self.app_settings.value("settings/out_folder_path", app_outdir),
        )
        # test selected folder
        if not path.exists(selected_folder):
            logger.error("No folder selected")
            return False
        else:
            selected_folder = path.realpath(selected_folder)
            logger.debug("Output folder selected: {}".format(selected_folder))

        # fill label and tooltip
        self.ui.lbl_output_folder_value.setText(path.basename(selected_folder))
        self.ui.lbl_output_folder_value.setToolTip(path.dirname(selected_folder))
        # save in settings
        self.app_settings.setValue(
            "settings/out_folder_label", path.basename(selected_folder)
        )
        self.app_settings.setValue("settings/out_folder_path", selected_folder)

        # connect the open button
        self.ui.btn_open_output_folder.pressed.disconnect()
        self.ui.btn_open_output_folder.pressed.connect(
            partial(self.app_utils.open_dir_file, selected_folder)
        )

    def settings_reset(self):
        """Reset settings to factiry defaults. Do not not remove authentication
        credentials. See #41
        """
        reset_msgbox = QMessageBox.question(
            self,
            self.tr("Settings - Reset to factory defaults"),
            self.tr(
                "Settings will be reinitialized (not"
                " authentication credentials).\n"
                "application will be closed."
            ),
            QMessageBox.Yes | QMessageBox.No
        )

        # close only if the user clicked yes
        if reset_msgbox == QMessageBox.Yes:
            logger.info("Settings - Reset to factory defaults.")
            self.app_settings.remove("formats")
            self.app_settings.remove("settings")
            self.settings_noSave = 1
            self.close()

    # -- UI Slots -------------------------------------------------------------
    @pyqtSlot(str, str, bool)
    def fill_app_props(
        self, app_infos_retrieved: str = "", latest_online_version: str = "", opencatalog_warning: bool = 0
    ):
        """Get app properties and fillfull the share frame in settings tab.

        :param str app_infos_retrieved: application information to display into the settings tabs
        :param str latest_online_version: latest version retrieved from GitHub to compare with the actual
        """
        # fill settings tab text
        self.ui.txt_shares.setText(app_infos_retrieved)
        # compare version used and online
        try:
            if semver.compare(__version__, latest_online_version) < 0:
                logger.info("A newer version is available.")
                version_msg = self.tr("New version available. You can download it here: ") + "https://github.com/isogeo/isogeo-2-office/releases/latest"
                self.setWindowTitle(self.windowTitle() + " ! " + version_msg)
            else:
                logger.debug("Used version is up-to-date")
                version_msg = self.tr("Version is up-to-date.")
        except Exception:
            logger.error("Version comparison failed.")

        # notification
        self.tray_icon.show()
        self.tray_icon.showMessage(
            "Isogeo to Office",
            self.tr("Application information has been retrieved.") + " " + version_msg,
            QIcon("resources/favicon.png"),
            2000,
        )
        self.update_status_bar(
            prog_step=0,
            status_msg=self.tr("Application information has been retrieved"),
        )

        # if needed, inform the user about a missing OpenCatalog
        if opencatalog_warning:
            oc_msg = self.tr("OpenCatalog is missing in one share at least. Check the settings tab to identify which one and fix it.")
            self.tray_icon.showMessage(
                "Isogeo to Office",
                oc_msg,
                QIcon("resources/favicon.png"),
            )
            self.update_status_bar(
                prog_step=0,
                status_msg=oc_msg,
                color="orange"
                )


    @pyqtSlot()
    def update_credentials(self):
        """Executed after credentials have been updated.
        """
        api_mngr.manage_api_initialization()
        self.init_api_connection()

    @pyqtSlot(MetadataSearch)
    def update_search_form(self, search: MetadataSearch):
        """Update search form with tags.

        :param MetadataSearch search: search returned with tags as dicts to update the search form.
        """
        # check if search returned some results
        if not search.total:
            logger.debug("The search returned no results. Please try other filters.")
            # stop progress bar and enable search form
            self.processing("end")
            self.ui.pgb_exports.setRange(0, 1)
            # inform user
            self.update_status_bar(prog_step=0, status_msg=self.tr("No results found. Please, try other filters."))
            # export button
            self.ui.btn_launch_export.setText(
                self.tr("Export {} metadata").format(search.total)
                )
            self.ui.btn_launch_export.setEnabled(False)
            return

        # get available search tags
        search_tags = search.tags
        logger.debug("Search tags keys: {}".format(list(search_tags)))

        # -- FILL FILTERS COMBOBOXES ------------------------------------------
        # clear previous state
        for cbb in self.cbbs_filters:
            cbb.clear()

        # add none selection item
        for cbb in self.cbbs_filters:
            cbb.addItem(" - ", "")

        # Add keywords
        for k, v in search_tags.get("keywords").items():
            self.ui.cbb_keyword.addItem(k, v)
        # Add owners
        for k, v in search_tags.get("owners").items():
            self.ui.cbb_owner.addItem(k, v)
        # Add shares
        for k, v in search_tags.get("shares").items():
            self.ui.cbb_share.addItem(k, v)
        # Add types
        for k, v in search_tags.get("types").items():
            self.ui.cbb_type.addItem(k, v)

        # export button
        self.ui.btn_launch_export.setText(
            self.tr("Export {} metadata").format(search.total)
        )

        # -- RESTORE PREVIOUS SELECTED FILTERS -------------------------------
        if self.search_type != "reset":
            query_tags = search.query.get("_tags")
            logger.debug("Previous query: {}".format(query_tags))
            # keywords
            if query_tags.get("keywords"):
                prev_val = list(query_tags.get("keywords"))[0]
                prev_idx = self.ui.cbb_keyword.findText(prev_val)
                self.ui.cbb_keyword.setCurrentIndex(prev_idx)
            else:
                pass
            # owners
            if query_tags.get("owners"):
                prev_val = list(query_tags.get("owners"))[0]
                prev_idx = self.ui.cbb_owner.findText(prev_val)
                self.ui.cbb_owner.setCurrentIndex(prev_idx)
            else:
                pass
            # shares
            if query_tags.get("shares"):
                # special case because share id aren't contextualized by the API
                #  so, the combobox is cleared and fillulled again with only the
                #  selected share and the non selector
                self.ui.cbb_share.clear()
                self.ui.cbb_share.addItem(" - ", "")
                # add selected share
                prev_val, prev_code = list(query_tags.get("shares").items())[0]
                self.ui.cbb_share.addItem(prev_val, prev_code)
                prev_idx = self.ui.cbb_share.findText(prev_val)
                self.ui.cbb_share.setCurrentIndex(prev_idx)
            else:
                pass
            # types
            if query_tags.get("types"):
                prev_val = list(query_tags.get("types"))[0]
                prev_idx = self.ui.cbb_type.findText(prev_val)
                self.ui.cbb_type.setCurrentIndex(prev_idx)
            else:
                pass
        else:
            pass

        # stop progress bar and enable search form
        self.processing("end", progbar_max=search.total)
        self.update_status_bar(prog_step=0, status_msg=self.tr("Search form updated"))

    @pyqtSlot(int, str)
    def update_status_bar(self, prog_step: int = 1, status_msg: str = "", duration: int = 0, color: str = None):
        """Display message into status bar.

        :param int prog_step: step to increase the progress bar. Defaults to 1.
        :param str status_msg: message to display into the status bar
        :param int duration: duration of the message in milliseconds.
        :param str color: color to apply to the message
        """
        # custom message foreground color
        if color is not None:
            self.ui.lbl_statusbar.setStyleSheet("color: {}".format(color))
        else:
            self.ui.lbl_statusbar.setStyleSheet("")
        # status bar and systray
        self.ui.lbl_statusbar.showMessage(status_msg, msecs=duration)
        self.tray_icon.setToolTip(status_msg)
        # progressbar
        prog_val = self.ui.pgb_exports.value() + prog_step
        self.ui.pgb_exports.setValue(prog_val)
        # check if progression is over
        if self.ui.pgb_exports.maximum() == prog_val and prog_step > 0:
            # display main window - see #22
            self.tray_icon.act_show.trigger()
            # notify the user - see #12
            self.tray_icon.showMessage(
                "Isogeo to Office",
                self.tr("Export operations are over."),
                QIcon("resources/favicon.png"),
                2000,
            )
            # open output dir - see #27
            self.app_utils.open_dir_file(
                self.app_settings.value("settings/out_folder_path", app_outdir)
            )


# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == "__main__":
    import sys

    # create the application and the main window
    app = QApplication(sys.argv)
    app.setOrganizationName("Isogeo")
    app.setOrganizationDomain("isogeo.com")
    app.setApplicationName("Isogeo To Office")
    app.setApplicationVersion(__version__)
    # apply dark style
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    # apply language
    locale_path = path.join(
        app_dir, "i18n", "IsogeoToOffice_{}.qm".format(current_locale.system().name())
    )
    translator = QTranslator()
    translator.load(path.realpath(locale_path))
    app.installTranslator(translator)
    # link to Isogeo to Office main UI
    i2o = IsogeoToOffice_Main()
    sys.exit(app.exec_())
