# -*- coding: UTF-8 -*-
#! python3

"""
    Isogeo To Office - Threads used to subprocess some tasks

    Author: Julien Moura (@geojulien) for Isogeo
    Python: 3.7.x
"""

# #############################################################################
# ########## Libraries #############
# ##################################

# standard library
import logging
from os import path, walk
from pathlib import Path
from tempfile import mkdtemp
from zipfile import ZipFile

# 3rd party library
from docxtpl import DocxTemplate
from isogeo_pysdk.models import Metadata, MetadataSearch, Share
from isogeotoxlsx import Isogeo2xlsx
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.cell import WriteOnlyCell
from PyQt5.QtCore import QDateTime, QLocale, QThread, pyqtSignal
import requests

# submodules - export
from . import Isogeo2docx, isogeo2office_utils

# #############################################################################
# ########## Globals ###############
# ##################################

app_utils = isogeo2office_utils()
current_locale = QLocale()
logger = logging.getLogger("isogeo2office")


# #############################################################################
# ######## QThreads ################
# ##################################

# API REQUESTS ----------------------------------------------------------------
class ThreadAppProperties(QThread):
    # signals
    sig_finished = pyqtSignal(str, str, bool)

    def __init__(self, api_manager: object):
        QThread.__init__(self)
        self.api_mngr = api_manager

    # run method gets called when we start the thread
    def run(self):
        """Get application informations and build the text to display into the settings tab.
        """
        # local vers
        opencatalog_warning = 0

        # insert text
        text = "<html>"  # opening html content
        # properties of the authenticated application
        app = self.api_mngr.isogeo.app_properties
        text += "<p>{}<a href='{}' style='color: CornflowerBlue;'>{}</a> ".format(
            self.tr("This application is authenticated as "),
            app.url or "http://help.isogeo.com/isogeo2office/",
            app.name or "Isogeo to Office",
        )
        logger.info("Application authenticated: {}".format(app.name))
        # shares feeding the application
        if len(self.api_mngr.isogeo._shares) == 1:
            text += "{}{} {}</p></br>".format(
                self.tr(" and powered by "), "1", self.tr("share:")
            )
        else:
            text += "{}{} {}</p></br>".format(
                self.tr(" and powered by "),
                len(self.api_mngr.isogeo._shares),
                self.tr("shares:"),
            )
        # shares details
        for s in self.api_mngr.isogeo._shares:
            share = Share(**s)
            # share variables
            creator_name = share._creator.get("contact").get("name", "")
            creator_email = share._creator.get("contact").get("email", "")

            # share administration URL
            text += "<p><a href='{}' style='color: CornflowerBlue;'><b>{}</b></a></p>".format(
                share.admin_url(self.api_mngr.isogeo.app_url), share.name
            )

            # OpenCatalog status - ref: https://github.com/isogeo/isogeo-2-office/issues/54
            opencatalog_url = share.opencatalog_url(self.api_mngr.isogeo.oc_url)
            if self.api_mngr.isogeo.head(opencatalog_url):
                text += "<p>{} <a href='{}' style='color: CornflowerBlue;'><b>{}</b></a></p>".format(
                    self.tr("OpenCatalog status:"), opencatalog_url, self.tr("enabled")
                )
            else:
                text += "<p>{} <span style='color: red;'>{}</span</p>".format(
                    self.tr("OpenCatalog status:"), self.tr("disabled")
                )
                opencatalog_warning = 1

            # last modification (share renamed, changes in catalogs or applications, etc.)
            text += "<p>{} {}</p>".format(
                self.tr("Updated:"),
                QDateTime(app_utils.hlpr_datetimes(share._modified)).toString(),
            )

            # workgroup contact owner of the share
            text += "<p>{} <a href='mailto:{}'>{}</a></p>".format(
                self.tr("Contact:"), creator_email, creator_name
            )
            text += "<p><hr></p>"
        text += "</html>"

        # -- Check IsogeoToOffice version ----------------
        proxies = self.api_mngr.isogeo.proxies
        # get latest release on Github
        try:
            latest_v = requests.get(
                "https://api.github.com/repos/isogeo/isogeo-2-office/releases?per_page=1",
                proxies=proxies,
            ).json()[0]
            online_version = latest_v.get("tag_name")
        except Exception as e:
            logger.error(
                "Unable to get the latest application version from Github: ".format(e)
            )
            online_version = "0.0.0"

        # handle version label starting with a non digit char
        if not online_version[0].isdigit():
            online_version = online_version[1:]

        # Now inform the main thread with the output (fill_app_props)
        self.sig_finished.emit(text, online_version, opencatalog_warning)


class ThreadSearch(QThread):
    # signals
    sig_finished = pyqtSignal(MetadataSearch, name="IsogeoSearch")

    def __init__(self, api_manager: object):
        QThread.__init__(self)
        self.api_mngr = api_manager
        self.search_params = dict

    # run method gets called when we start the thread
    def run(self):
        """Get application and informations
        """
        logger.debug("Search started.")
        search = self.api_mngr.isogeo.search(**self.search_params)
        logger.debug(
            "Search finished: {} results on {} total. Transmitting to slot...".format(
                len(search.results), search.total
            )
        )
        # Search request finished
        self.sig_finished.emit(search)


# EXPORTS ---------------------------------------------------------------------
class ThreadExportExcel(QThread):
    # signals
    sig_step = pyqtSignal(int, str, name="ExportExcel")

    def __init__(
        self,
        search_to_export: MetadataSearch,
        output_path: str = r"output/",
        url_base_edit: str = "https://app.isogeo.com/",
        url_base_view: str = "https://open.isogeo.com/",
        shares: list = [],
        opt_attributes: int = 0,
        opt_dasboard: int = 0,
        opt_fillfull: int = 0,
        opt_inspire: int = 0,
    ):
        QThread.__init__(self)
        # export settings
        self.search = search_to_export
        self.output_xlsx_path = output_path
        self.url_base_edit = url_base_edit
        self.url_base_view = url_base_view
        self.shares = shares
        self.opt_attributes = opt_attributes
        self.opt_dasboard = opt_dasboard
        self.opt_fillfull = opt_fillfull
        self.opt_inspire = opt_inspire

    # run method gets called when we start the thread
    def run(self):
        """Export metadata into an Excel workbook
        """
        language = current_locale.name()[:2]
        # workbook
        wb = Isogeo2xlsx(
            lang=language,
            url_base_edit=self.url_base_edit,
            url_base_view=self.url_base_view,
        )
        # worksheets
        wb.set_worksheets(
            auto=self.search.tags.keys(),
            dashboard=self.opt_dasboard,
            attributes=self.opt_attributes,
            fillfull=self.opt_fillfull,
            inspire=self.opt_inspire,
        )

        # parsing metadata
        for md in self.search.results:
            # load metadata
            metadata = Metadata.clean_attributes(md)
            # show progression
            self.sig_step.emit(
                1, self.tr("Processing Excel: {}").format(metadata.title_or_name())
            )

            # opencatalog url - get the matching share
            matching_share = app_utils.get_matching_share(
                metadata=metadata, shares=self.shares
            )

            # store metadata
            wb.store_metadatas(metadata, share=matching_share)

        # complementary analisis
        self.sig_step.emit(
            0,
            self.tr("Processing Excel: {}").format(
                self.tr("complementary analisis...")
            ),
        )
        wb.launch_analisis()

        # tunning full worksheet
        self.sig_step.emit(
            0, self.tr("Processing Excel: {}").format(self.tr("tunning sheets..."))
        )
        wb.tunning_worksheets()

        # save workbook
        try:
            wb.save(self.output_xlsx_path)
        except PermissionError as e:
            logger.error(e)
            wb.close()
            wb.save(path.normpath(self.output_xlsx_path))

        # Excel export finished
        # Now inform the main thread with the output (fill_app_props)
        logger.info("Excel - Export is over")
        self.sig_step.emit(0, self.tr("Excel finished"))
        # self.deleteLater()


class ThreadExportWord(QThread):
    """QThread used to export an Isogeo search into metadata.

    :param MetadataSearch search_to_export: metadata to dumpinto the template
    :param str output_path: path to the output folder to store the generated Word
    :param str tpl_path: path to the Word template to use
    :param str url_base_edit: base url to format edit links. Defaults to: https://app.isogeo.com
    :param str url_base_view: base url to format view links. Defaults to: https://open.isogeo.com
    :param list shares: list of shares feeding the application
    :param dict thumbnails: matching table between metadata and image path
    :param str timestamp: timestamp used to name the output file
    :param int length_uuid: number of UUID characters to use to name the output file
    """

    # signals
    sig_step = pyqtSignal(int, str, name="ExportWord")

    def __init__(
        self,
        search_to_export: MetadataSearch = {},
        output_path: str = r"output/",
        tpl_path: str = r"templates/template_Isogeo.docx",
        url_base_edit: str = "https://app.isogeo.com/",
        url_base_view: str = "https://open.isogeo.com/",
        shares: list = [],
        thumbnails: dict = {},
        timestamp: str = "",
        length_uuid: int = 0,
    ):
        QThread.__init__(self)
        # export settings
        self.search = search_to_export
        self.output_docx_folder = output_path
        self.url_base_edit = url_base_edit
        self.url_base_view = url_base_view
        self.shares = shares
        self.tpl_path = path.realpath(tpl_path)
        self.thumbnails = thumbnails
        self.timestamp = timestamp
        self.length_uuid = length_uuid

    # run method gets called when we start the thread
    def run(self):
        """Export each metadata from a search results into a Word document."""
        # vars
        language = current_locale.name()[:2]
        thumbnail_default = ("", path.realpath(r"resources/favicon.png"))

        # word generator
        to_docx = Isogeo2docx(
            lang=language,
            url_base_edit=self.url_base_edit,
            url_base_view=self.url_base_view,
        )

        # parsing metadata
        for md in self.search.results:
            # load metadata
            metadata = Metadata.clean_attributes(md)

            # progression
            self.sig_step.emit(
                1, self.tr("Processing Word: {}").format(metadata.title_or_name())
            )

            # opencatalog url - get the matching share
            matching_share = app_utils.get_matching_share(
                metadata=metadata, shares=self.shares
            )

            # thumbnails
            thumbnail_abs_path = self.thumbnails.get(metadata._id, thumbnail_default)[1]
            if not thumbnail_abs_path or not path.isfile(thumbnail_abs_path):
                thumbnail_abs_path = path.realpath(r"resources/favicon.png")
            logger.debug("Thumbnail used: {}".format(thumbnail_abs_path))
            metadata.thumbnail = thumbnail_abs_path

            # templating
            tpl = DocxTemplate(self.tpl_path)
            # fill template
            to_docx.md2docx(docx_template=tpl, md=metadata, share=matching_share)
            # filename
            md_name = metadata.title_or_name(slugged=1)
            uuid = "{}".format(metadata._id[: self.length_uuid])

            out_docx_filename = "{}_{}_{}.docx".format(
                self.output_docx_folder, md_name, uuid
            )
            # saving
            logger.debug("Saving output Word docx: {}".format(out_docx_filename))
            try:
                tpl.save(out_docx_filename)
            except Exception as e:
                logger.error(e)
                self.sig_step.emit(
                    0,
                    self.tr("Word: error occurred during saving step. Check the log."),
                )
            del tpl

        # Word export finished
        # Now inform the main thread with the output (fill_app_props)
        logger.info("Word - Export is over")
        self.sig_step.emit(0, self.tr("Word finished"))
        # self.deleteLater()


class ThreadExportXml(QThread):
    # signals
    sig_step = pyqtSignal(int, str, name="ExportXML")

    def __init__(
        self,
        search_to_export: dict = {},
        isogeo_api_mngr: object = None,
        output_path: str = r"output/",
        opt_zip: int = 0,
        timestamp: str = "",
        length_uuid: int = 0,
    ):
        QThread.__init__(self)
        # export settings
        self.search = search_to_export
        self.api_mngr = isogeo_api_mngr
        self.output_xml_path = output_path
        self.opt_zip = opt_zip
        self.timestamp = timestamp
        self.length_uuid = length_uuid

    # run method gets called when we start the thread
    def run(self):
        """Export each metadata into an XML ISO 19139
        """
        # ZIP or not ZIP
        if self.opt_zip:
            # into a temporary directory
            out_dir = mkdtemp(prefix="IsogeoToOffice_", suffix="_xml")
            logger.debug("XML - Temporary directory created: {}".format(out_dir))
        else:
            # directly into the output directory
            out_dir = path.realpath(self.output_xml_path)

        # parsing metadata
        for md in self.search.results:
            # load metadata
            metadata = Metadata.clean_attributes(md)
            # progression
            self.sig_step.emit(
                1, self.tr("Processing XML: {}").format(metadata.title_or_name())
            )

            # filename
            md_name = metadata.title_or_name(slugged=1)
            uuid = "{}".format(md.get("_id")[: self.length_uuid])

            # compressed or raw
            if self.opt_zip:
                out_xml_path = path.join(out_dir, "{}_{}.xml".format(md_name, uuid))
            else:
                out_xml_path = out_dir + "_{}_{}.xml".format(md_name, uuid)
            logger.debug("XML - Output path: {}".format(out_xml_path))

            # export
            xml_stream = self.api_mngr.isogeo.metadata.download_xml(metadata)
            with open(path.realpath(out_xml_path), "wb") as out_md:
                for block in xml_stream.iter_content(1024):
                    out_md.write(block)

        # ZIP or not ZIP
        if self.opt_zip:
            out_zip_path = self.output_xml_path + ".zip"
            final_zip = ZipFile(out_zip_path, "w")
            for root, dirs, files in walk(out_dir):
                for f in files:
                    final_zip.write(path.join(root, f), f)
            final_zip.close()
            logger.debug("XML - ZIP: {}".format(out_zip_path))
        else:
            pass

        # XML export finished
        # Now inform the main thread with the output (fill_app_props)
        logger.info("XML - Export is over")
        self.sig_step.emit(0, "XML finished")
        # self.deleteLater()


# TO IMPORT LATER -------------------------------------------------------------
class ThreadThumbnails(QThread):
    # signals
    sig_step = pyqtSignal(int, str)

    def __init__(
        self,
        search_to_export: dict = {},
        output_path: str = r"thumbnails/thumbnails.xlsx",
        thumbnails: dict = {},
    ):
        QThread.__init__(self)
        # export settings
        self.search = search_to_export
        self.output_xlsx_path = output_path
        self.thumbnails = thumbnails

    # run method gets called when we start the thread
    def run(self):
        """Build thumbnail matchign structure into an Excel workbook

        1. match with existing thumbnails
        2. write new file
        3. archive previous files
        """
        # workbook structure
        wb = Workbook(write_only=True)
        ws = wb.create_sheet("i2o_thumbnails")
        # columns dimensions
        ws.column_dimensions["A"].width = 35
        ws.column_dimensions["B"].width = 75
        ws.column_dimensions["C"].width = 75
        # headers values
        head_col1 = WriteOnlyCell(ws, value="isogeo_uuid")
        head_col2 = WriteOnlyCell(ws, value="isogeo_title_slugged")
        head_col3 = WriteOnlyCell(ws, value="img_abs_path")
        # headers comments
        comment = Comment(text="Do not modify worksheet structure", author="Isogeo")

        head_col1.comment = head_col2.comment = head_col3.comment = comment

        # headers styling
        head_col1.style = head_col2.style = head_col3.style = "Headline 2"
        # insert headers
        ws.append((head_col1, head_col2, head_col3))

        # parsing metadata
        li_exported_md = []
        for md in self.search.results:
            # show progression
            md_title = md.get("title", "No title")
            self.sig_step.emit(
                0, self.tr("Preparing thumbnail table for: {}").format(md_title)
            )
            # thumbnail matching
            thumbnail_abs_path = self.thumbnails.get(md.get("_id"), "  ")[1]
            # fill with metadata
            ws.append(
                (
                    md.get("_id"),
                    app_utils.clean_filename(
                        md.get("title", md.get("name", "NR")), mode="strict"
                    ),
                    thumbnail_abs_path,
                )
            )
            # list exported metadata to compare with previous
            li_exported_md.append(md.get("_id"))

        # insert previous metadata which have not been exported this time
        for thumb, title_path in self.thumbnails.items():
            if thumb not in li_exported_md:
                try:
                    ws.append((thumb, title_path[0], title_path[1]))
                except TypeError as e:
                    logger.error("Thumbnails table error: {}".format(e))

        # save workbook
        try:
            wb.save(self.output_xlsx_path)
        except PermissionError as e:
            logger.error(e)
            wb.close()
            wb.save(path.normpath(self.output_xlsx_path))

        # Excel export finished
        # Now inform the main thread with the output (fill_app_props)
        logger.info("Thumbnail - Table creation is over")
        self.sig_step.emit(0, self.tr("Thumbnail table finished"))
        # self.deleteLater()


class ThreadExportHtmlReport(QThread):
    # signals
    sig_step = pyqtSignal(int, str)

    def __init__(
        self,
        search_to_export: dict = {},
        isogeo_api_mngr: object = None,
        html_template: str = "",
        output_path: str = r"output/",
        timestamp: str = "",
    ):
        QThread.__init__(self)
        # export settings
        self.search = search_to_export
        self.api_mngr = isogeo_api_mngr
        self.html_tpl = html_template
        self.output_path = output_path
        self.timestamp = timestamp

    # run method gets called when we start the thread
    def run(self):
        """Export each metadata into an XML ISO 19139
        """
        logger.debug("hohohoho")
        # ZIP or not ZIP
        stats_types = [100, 50, 25, 15]

        # output template
        out_html_file = self.html_tpl.render(
            varTitle="Isogeo To Office - Report", varLblChartTypes="Metadata by types"
        )

        with open("test_out_html_templated.html", "w") as fh:
            fh.write(out_html_file)


# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == "__main__":
    logging.critical("Can't be used as main script.")
