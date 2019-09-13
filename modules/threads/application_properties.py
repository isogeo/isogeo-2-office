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
from isogeo_pysdk.models import Share
from PyQt5.QtCore import QDateTime, QLocale, QThread, pyqtSignal
import requests

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

        # check if the application is not connected to any share
        if app is None:
            logger.warning("Application has no shares to retrieve. It can't work.")
            text += "<p><span style='color: red;'>{}</span></p></html>".format(
                self.tr("No share is connected to the application. It cannot work.")
            )
            online_version = "0.0.0"
            opencatalog_warning = 1
            self.sig_finished.emit(text, online_version, opencatalog_warning)
            return

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
                text += "<p>{} <span style='color: red;'>{}</span></p>".format(
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
                "Unable to get the latest application version from Github: {}".format(e)
            )
            online_version = "0.0.0"

        # handle version label starting with a non digit char
        if not online_version[0].isdigit():
            online_version = online_version[1:]

        # Now inform the main thread with the output (fill_app_props)
        self.sig_finished.emit(text, online_version, opencatalog_warning)


# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == "__main__":
    logging.critical("Can't be used as main script.")
