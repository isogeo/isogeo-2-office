# -*- coding: UTF-8 -*-
#! python3  # noqa: E265

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
from isogeo_pysdk.models import Metadata, MetadataSearch, Share
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.cell import WriteOnlyCell
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


# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == "__main__":
    logging.critical("Can't be used as main script.")
