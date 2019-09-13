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
from isogeo_pysdk.models import Metadata, MetadataSearch
from isogeotoxlsx import Isogeo2xlsx
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


# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == "__main__":
    logging.critical("Can't be used as main script.")
