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

# 3rd party library
from isogeo_pysdk.models import MetadataSearch
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

# API REQUESTS ----------------------------------------------------------------
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


# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == "__main__":
    logging.critical("Can't be used as main script.")
