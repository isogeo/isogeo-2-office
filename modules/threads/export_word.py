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
from os import path

# 3rd party library
from docxtpl import DocxTemplate
from isogeo_pysdk.models import Metadata, MetadataSearch
from isogeotodocx import Isogeo2docx
from PyQt5.QtCore import QLocale, QThread, pyqtSignal

# submodules - export
from modules.utils import isogeo2office_utils

# #############################################################################
# ########## Globals ###############
# ##################################

app_utils = isogeo2office_utils()
current_locale = QLocale()
logger = logging.getLogger("isogeo2office")


# #############################################################################
# ######## QThreads ################
# ##################################

# EXPORTS ---------------------------------------------------------------------
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

        # word generator
        to_docx = Isogeo2docx(
            lang=language,
            thumbnails=self.thumbnails,
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


# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == "__main__":
    logging.critical("Can't be used as main script.")
