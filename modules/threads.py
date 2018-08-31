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
from logging.handlers import RotatingFileHandler
from os import path

# 3rd party library
from docxtpl import DocxTemplate
from PyQt5.QtCore import (QDate, QLocale, QThread, pyqtSignal, pyqtSlot)

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
            text += "{}{}{}</p></br>".format(self.tr(" powered by "),
                                             "1",
                                             self.tr("share:"))
        else:
            text += "{}{}{}</p></br>".format(self.tr(" powered by "),
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
        self.signal.emit(text)


class ExportExcelThread(QThread):
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
        # workbook
        wb = Isogeo2xlsx(lang="FR",
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
            self.sig_step.emit(1, "Processing Excel: {}".format(md_title))
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
        self.sig_step.emit(0, "Excel finished")


class ExportWordThread(QThread):
    # signals
    sig_step = pyqtSignal(int, str)

    def __init__(self,
                 search_to_export: dict = {},
                 output_path: str = r"output/",
                 tpl_path: str = r"templates/template_Isogeo.docx",
                 timestamp: str = "",
                 length_uuid: int = 0):
        QThread.__init__(self)
        # export settings
        self.search = search_to_export
        self.output_docx_folder = output_path
        self.tpl_path = path.realpath(tpl_path)
        self.timestamp = timestamp
        self.length_uuid = length_uuid

    # run method gets called when we start the thread
    def run(self):
        """Export each metadata into a Word document
        """
        # word generator
        to_docx = Isogeo2docx()

        # parsing metadata
        for md in self.search.get("results"):
            # progression
            md_title = md.get("title", "No title")
            self.sig_step.emit(1, "Processing Word: {}".format(md_title))
            # templating
            tpl = DocxTemplate(self.tpl_path)
            # fill template
            to_docx.md2docx(docx_template=tpl,
                            md=md,
                            url_base="https://open.isogeo.com",
                            thumb_path="https://www.isogeo.com/images/logo.png")
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
                self.sig_step.emit(0, "Word: Error")
            del tpl

        # Word export finished
        # Now inform the main thread with the output (fill_app_props)
        self.sig_step.emit(0, "Word finished")



# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == "__main__":
    logging.critical("Can't be used as main script.")
