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
from os import path, walk
from tempfile import mkdtemp
from zipfile import ZipFile

# 3rd party library
from isogeo_pysdk.models import Metadata
from PyQt5 import QtCore
import requests

# submodules - export
from modules.utils import isogeo2office_utils

# #############################################################################
# ########## Globals ###############
# ##################################

app_utils = isogeo2office_utils()
current_locale = QtCore.QLocale()
logger = logging.getLogger("isogeo2office")


# #############################################################################
# ######## QThreads ################
# ##################################

# EXPORTS ---------------------------------------------------------------------
class ThreadExportXml(QtCore.QThread):
    # signals
    sig_step = QtCore.pyqtSignal(int, str, name="ExportXML")

    def __init__(
        self,
        search_to_export: dict = {},
        isogeo_api_mngr: object = None,
        output_path: str = r"output/",
        opt_zip: int = 0,
        timestamp: str = "",
        length_uuid: int = 0,
    ):
        QtCore.QThread.__init__(self)
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
            try:
                xml_stream = self.api_mngr.isogeo.metadata.download_xml(metadata)
            except requests.Timeout as e:
                logger.error(
                    "Connection with Isogeo failed, trying to get the XML version of metadata: {} ({})."
                    " Original error: {}".format(
                        metadata.title_or_name(slugged=1), metadata._id, e
                    )
                )
                continue

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


# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == "__main__":
    logging.critical("Can't be used as main script.")
