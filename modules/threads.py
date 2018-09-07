# -*- coding: UTF-8 -*-
#! python3

"""
    Isogeo To Office - Threads used to subprocess some tasks

    Author: Julien Moura (@geojulien)
    Python: 3.6.x
"""

# #############################################################################
# ########## Libraries #############
# ##################################

# standard library
import logging
from os import path, walk
from tempfile import mkdtemp
from zipfile import ZipFile

# 3rd party library
from docxtpl import DocxTemplate
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.worksheet.write_only import WriteOnlyCell
from PyQt5.QtCore import QDate, QLocale, QThread, pyqtSignal

# submodules - export
from . import Isogeo2docx, Isogeo2xlsx, isogeo2office_utils

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
    sig_finished = pyqtSignal(str)

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
        text += "<p>{}<a href='{}' style='color: CornflowerBlue;'>{}</a> "\
                .format(self.tr("This application is authenticated as "),
                        app.get("url", "https://isogeo.gitbooks.io/app-isogeo2office/content/"),
                        app.get("name", "Isogeo to Office"))
        # shares feeding the application
        if len(shares) == 1:
            text += "{}{} {}</p></br>".format(self.tr(" and powered by "),
                                              "1",
                                              self.tr("share:"))
        else:
            text += "{}{} {}</p></br>".format(self.tr(" and powered by "),
                                              len(shares),
                                              self.tr("shares:"))
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
        self.sig_finished.emit(text)


class ThreadSearch(QThread):
    # signals
    sig_finished = pyqtSignal(dict)

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
        logger.debug("Search finished.")
        # Search request finished
        self.sig_finished.emit(search)


# EXPORTS ---------------------------------------------------------------------
class ThreadExportExcel(QThread):
    # signals
    sig_step = pyqtSignal(int, str)

    def __init__(self,
                 search_to_export: dict = {},
                 output_path: str = r"output/",
                 opt_attributes: int = 1,
                 opt_dasboard: int = 1,
                 opt_fillfull: int = 1,
                 opt_inspire: int = 1):
        QThread.__init__(self)
        # export settings
        self.search = search_to_export
        self.output_xlsx_path = output_path
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
        wb = Isogeo2xlsx(lang=language,
                         url_base="https://open.isogeo.com")
        wb.set_worksheets(auto=self.search.get('tags').keys(),
                          dashboard=self.opt_dasboard,
                          attributes=self.opt_attributes,
                          fillfull=self.opt_fillfull,
                          inspire=self.opt_inspire)

        # parsing metadata
        for md in self.search.get("results"):
            # show progression
            md_title = md.get("title", "No title")
            self.sig_step.emit(1, self.tr("Processing Excel: {}").format(md_title))
            # store metadata
            wb.store_metadatas(md)

        # tunning full worksheet
        wb.tunning_worksheets()

        # special sheets
        if self.opt_attributes:
            wb.analisis_attributes()
        else:
            pass

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
        self.deleteLater()


class ThreadExportWord(QThread):
    # signals
    sig_step = pyqtSignal(int, str)

    def __init__(self,
                 search_to_export: dict = {},
                 output_path: str = r"output/",
                 tpl_path: str = r"templates/template_Isogeo.docx",
                 thumbnails: dict = {},
                 timestamp: str = "",
                 length_uuid: int = 0):
        QThread.__init__(self)
        # export settings
        self.search = search_to_export
        self.output_docx_folder = output_path
        self.tpl_path = path.realpath(tpl_path)
        self.thumbnails = thumbnails
        self.timestamp = timestamp
        self.length_uuid = length_uuid

    # run method gets called when we start the thread
    def run(self):
        """Export each metadata into a Word document
        """
        # vars
        thumbnail_default = ("", path.realpath(r"resources/credits/isogeo.png"))

        # word generator
        to_docx = Isogeo2docx()

        # parsing metadata
        for md in self.search.get("results"):
            # progression
            md_title = md.get("title", "No title")
            self.sig_step.emit(1, self.tr("Processing Word: {}").format(md_title))
            # thumbnails
            thumbnail_abs_path = self.thumbnails.get(md.get("_id"),
                                                     thumbnail_default)[1]
            logger.debug("Thumbnail used: {}".format(thumbnail_abs_path))
            md["thumbnail_local"] = thumbnail_abs_path
            # templating
            tpl = DocxTemplate(self.tpl_path)
            # fill template
            to_docx.md2docx(docx_template=tpl,
                            md=md,
                            url_base="https://open.isogeo.com")
            # filename
            md_name = app_utils.clean_filename(md.get("name",
                                                      md.get("title", "NR"))
                                               )
            if '.' in md_name:
                md_name = md_name.split(".")[1]
            else:
                pass
            uuid = "{}".format(md.get("_id")[:self.length_uuid])

            out_docx_filename = "{}_{}_{}.docx".format(self.output_docx_folder,
                                                       md_name,
                                                       uuid)
            # saving
            logger.debug("Saving output Word docx: {}".format(out_docx_filename))
            try:
                tpl.save(out_docx_filename)
            except Exception as e:
                logger.error(e)
                self.sig_step.emit(0, self.tr("Word: error occurred during saving step. Check the log."))
            del tpl

        # Word export finished
        # Now inform the main thread with the output (fill_app_props)
        logger.info("Word - Export is over")
        self.sig_step.emit(0, self.tr("Word finished"))
        self.deleteLater()


class ThreadExportXml(QThread):
    # signals
    sig_step = pyqtSignal(int, str)

    def __init__(self,
                 search_to_export: dict = {},
                 isogeo_api_mngr: object = None,
                 output_path: str = r"output/",
                 opt_zip: int = 0,
                 timestamp: str = "",
                 length_uuid: int = 0):
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
        for md in self.search.get("results"):
            # progression
            md_title = md.get("title", "No title")
            self.sig_step.emit(1, self.tr("Processing XML: {}").format(md_title))

            # filename
            md_name = app_utils.clean_filename(md.get("title",
                                                      md.get("name", "No name"))
                                               ).split(" -")[0]
            if '.' in md_name:
                md_name = md_name.split(".")[1]
            else:
                pass
            uuid = "{}".format(md.get("_id")[:self.length_uuid])

            if self.opt_zip:
                out_xml_path = path.join(out_dir, "{}_{}.xml".format(md_name, uuid))
            else:
                out_xml_path = out_dir + "_{}_{}.xml".format(md_name, uuid)
            logger.debug("XML - Output path: {}".format(out_xml_path))
            # export
            xml_stream = self.api_mngr.isogeo.xml19139(self.api_mngr.token,
                                                       md.get("_id"))
            with open(path.realpath(out_xml_path), 'wb') as out_md:
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
        self.deleteLater()


# TO IMPORT LATER -------------------------------------------------------------
class ThreadThumbnails(QThread):
    # signals
    sig_step = pyqtSignal(int, str)

    def __init__(self,
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
        comment = Comment(text="Do not modify worksheet structure",
                        author="Isogeo")

        head_col1.comment = head_col2.comment = head_col3.comment = comment

        # headers styling
        head_col1.style = head_col2.style = head_col3.style = "Headline 2"
        # insert headers
        ws.append((head_col1,
                   head_col2,
                   head_col3)
        )

        # parsing metadata
        li_exported_md = []
        for md in self.search.get("results"):
            # show progression
            md_title = md.get("title", "No title")
            self.sig_step.emit(1, self.tr("Preparing thumbnail table for: {}")
                                          .format(md_title))
            # thumbnail matching
            thumbnail_abs_path = self.thumbnails.get(md.get("_id"),
                                                     "  ")[1]
            # fill with metadata
            ws.append((md.get("_id"),
                       app_utils.clean_filename(md.get("title", md.get("name", "NR")),
                                                mode="strict"),
                       thumbnail_abs_path))
            # list exported metadata to compare with previous
            li_exported_md.append(md.get("_id"))

        # insert previous metadata which have not been exported this time
        for thumb, title_path in self.thumbnails.items():
            if thumb not in li_exported_md:
                ws.append((thumb,
                           title_path[0],
                           title_path[1]))

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
        self.sig_step.emit(0, self.tr("Thmbnail table finished"))
        self.deleteLater()


# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == "__main__":
    logging.critical("Can't be used as main script.")
