# -*- coding: UTF-8 -*-
#! python3

# Standard library
from functools import partial
import json
import logging
from os import path, rename
import time

# PyQT
from PyQt5.QtWidgets import QDialog, QErrorMessage
from PyQt5.QtCore import QLocale, QSettings

# Isogeo
from isogeo_pysdk import Isogeo, IsogeoChecker
from isogeo_pysdk import __version__ as pysdk_version

# submodules
from .utils import isogeo2office_utils

# ############################################################################
# ########## Globals ###############
# ##################################

app_utils = isogeo2office_utils()
current_locale = QLocale()
logger = logging.getLogger("isogeo2office")
qsettings = QSettings('Isogeo', 'IsogeoToOffice')

# ############################################################################
# ########## Classes ###############
# ##################################


class IsogeoApiMngr(object):
    """Isogeo API manager."""
    # Isogeo API wrapper
    isogeo = Isogeo
    token = str
    # ui reference - authentication form
    ui_auth_form = QDialog
    auth_form_request_url = "https://www.isogeo.com"
    # api parameters
    api_app_id = ""
    api_app_secret = ""
    api_url_base = "https://v1.api.isogeo.com/"
    api_url_auth = "https://id.api.isogeo.com/oauth/authorize"
    api_url_token = "https://id.api.isogeo.com/oauth/token"
    api_url_redirect = "http://localhost:5000/callback"

    # plugin credentials storage parameters
    credentials_storage = {"QSettings": 0,
                           "oAuth2_file": 0,
                           }
    auth_folder = ""

    # API URLs - Prod
    platform, api_url, app_url, csw_url,\
        mng_url, oc_url, ssl = app_utils.set_base_url("prod")

    def __init__(self):
        super(IsogeoApiMngr, self)

    # MANAGER -----------------------------------------------------------------
    def manage_api_initialization(self):
        """Perform several operations to use Isogeo API:

        1. check if existing credentials are stored somewhere
        2. check if credentials are valid requesting Isogeo API ID
        """
        # try to retrieve existing credentials from potential sources
        self.credentials_storage["QSettings"] = self.credentials_check_qsettings()
        self.credentials_storage["oAuth2_file"] = self.credentials_check_file()

        # update class attributes from credentials found
        if self.credentials_storage.get("QSettings"):
            self.credentials_update("QSettings")
        elif self.credentials_storage.get("oAuth2_file"):
            self.credentials_update("oAuth2_file")
        else:
            logger.info("No credentials found. ")
            self.display_auth_form()
            return False

        # start api wrapper
        try:
            self.isogeo = Isogeo(client_id=self.api_app_id,
                                 client_secret=self.api_app_secret,
                                 lang=current_locale.name()[:2])
            self.token = self.isogeo.connect()
            return True
        except ValueError as e:
            logger.error(e)
            self.display_auth_form()
        except Exception as e:
            logger.error(e)
            self.display_auth_form()

    # CREDENTIALS LOCATORS ----------------------------------------------------
    def credentials_check_qsettings(self):
        """Retrieve Isogeo API credentials within APP QSettings."""
        if "auth" in qsettings.childGroups() and qsettings.contains("auth/app_id"):
            logger.debug("Credentials found within QSettings: isogeo/")
            return True
        else:
            logger.debug("No Isogeo credentials found within QSettings.")
            return False

    def credentials_check_file(self):
        """Retrieve Isogeo API credentials from a file stored inside the
        plugin _auth subfolder.
        """
        credentials_filepath = path.join(self.auth_folder, "client_secrets.json")
        # check if a client_secrets.json fil is stored inside the _auth subfolder
        if not path.isfile(credentials_filepath):
            logger.debug("No credential files found: {}"
                         .format(credentials_filepath))
            return False
        # check file structure
        try:
            app_utils.credentials_loader(credentials_filepath)
            logger.debug("Credentials found in {}".format(credentials_filepath))
        except Exception as e:
            logger.debug(e)
            return False
        # end of method
        return True

    # CREDENTIALS SAVER -------------------------------------------------------
    def credentials_storer(self, store_location="QSettings"):
        """Store class credentials attributes into the specified store_location.

        :param store_location str: name of targetted store location. Options:
            - QSettings
        """
        if store_location == "QSettings":
            qsettings.setValue("auth/app_id", self.api_app_id)
            qsettings.setValue("auth/app_secret", self.api_app_secret)
            qsettings.setValue("auth/url_base", self.api_url_base)
            qsettings.setValue("auth/url_auth", self.api_url_auth)
            qsettings.setValue("auth/url_token", self.api_url_token)
            qsettings.setValue("auth/url_redirect", self.api_url_redirect)
        else:
            pass
        logger.debug("Credentials stored into: {}".format(store_location))

    def credentials_update(self, credentials_source="QSettings"):
        """Update class attributes from specified credentials source."""
        # update class attributes
        if credentials_source == "QSettings":
            self.api_app_id = qsettings.value("auth/app_id", "")
            self.api_app_secret = qsettings.value("auth/app_secret", "")
            self.api_url_base = qsettings.value("auth/url_base", "https://v1.api.isogeo.com/")
            self.api_url_auth = qsettings.value("auth/url_auth", "https://id.api.isogeo.com/oauth/authorize")
            self.api_url_token = qsettings.value("auth/url_token", "https://id.api.isogeo.com/oauth/token")
            self.api_url_redirect = qsettings.value("auth/url_redirect", "http://localhost:5000/callback")
        elif credentials_source == "oAuth2_file":
            creds = app_utils.credentials_loader(path.join(self.auth_folder,
                                                           "client_secrets.json"))
            self.api_app_id = creds.get("client_id")
            self.api_app_secret = creds.get("client_secret")
            self.api_url_base = creds.get("uri_base")
            self.api_url_auth = creds.get("uri_auth")
            self.api_url_token = creds.get("uri_token")
            self.api_url_redirect = creds.get("uri_redirect")
        else:
            pass

        logger.debug("Credentials updated from: {}. Application ID used: {}"
                     .format(credentials_source, self.api_app_id))

    # AUTHENTICATION FORM -----------------------------------------------------
    def display_auth_form(self):
        """Show authentication form with prefilled fields."""
        # connect widgets
        self.ui_auth_form.chb_isogeo_editor\
                         .stateChanged\
                         .connect(lambda: qsettings.setValue("user/editor",
                                                             int(self.ui_auth_form.chb_isogeo_editor.isChecked())
                                                             )
                                  )
        self.ui_auth_form.btn_ok_cancel.clicked.connect(self.ui_auth_form.close)
        # button to request an account by email
        self.ui_auth_form.btn_account_new.pressed.connect(
            partial(app_utils.open_urls,
                    [self.auth_form_request_url, ]))

        # fillfull auth form fields from stored settings
        self.ui_auth_form.btn_ok_cancel.setEnabled(0)
        self.ui_auth_form.ent_app_id.setText(self.api_app_id)
        self.ui_auth_form.ent_app_secret.setText(self.api_app_secret)
        self.ui_auth_form.lbl_api_url_value.setText(self.api_url_base)
        self.ui_auth_form.chb_isogeo_editor.setChecked(qsettings
                                                       .value("user/editor", 0))
        # display
        logger.debug("Authentication form filled and ready to be launched.")
        self.ui_auth_form.show()
        self.ui_auth_form.setFocus()

    def credentials_uploader(self):
        """Get file selected by the user and loads API credentials into plugin.
        If the selected is compliant, credentials are loaded from then it's
        moved inside plugin/_auth subfolder.
        """
        selected_file = app_utils.open_FileNameDialog(self.ui_auth_form)
        logger.debug("Credentials file picker (QFileDialog) returned: {}"
                     .format(selected_file))
        # test file structure
        if not path.exists(selected_file[0]):
            logger.error("No file selected")
            return False
        else:
            logger.debug("Selected file exists.")
        try:
            selected_file = path.normpath(selected_file[0])
            api_credentials = app_utils.credentials_loader(selected_file)
        except Exception as e:
            logger.error("Selected file is bad formatted: {}".format(e))
            return False

        # move credentials file into the plugin file structure
        if path.isfile(path.join(self.auth_folder, "client_secrets.json")):
            rename(path.join(self.auth_folder, "client_secrets.json"),
                   path.join(self.auth_folder, "old_client_secrets_{}.json"
                                               .format(int(time.time())))
                   )
            logger.debug("client_secrets.json already existed. "
                         "Previous file has been renamed.")
        else:
            pass
        rename(selected_file,
               path.join(self.auth_folder, "client_secrets.json"))
        logger.debug("Selected credentials file has been moved into plugin"
                     "_auth subfolder")

        # check validity
        try:
            self.isogeo = Isogeo(client_id=api_credentials.get("client_id"),
                                 client_secret=api_credentials.get("client_secret")
                                 )
        except Exception as e:
            logger.debug(e)
            return False

        # set form
        self.ui_auth_form.ent_app_id.setText(api_credentials.get("client_id"))
        self.ui_auth_form.ent_app_secret.setText(api_credentials.get("client_secret"))
        self.ui_auth_form.lbl_api_url_value.setText(api_credentials.get("uri_auth"))
        self.ui_auth_form.btn_ok_cancel.setEnabled(1)

        # update class attributes from file
        self.credentials_update(credentials_source="oAuth2_file")

        # store into QSettings if existing
        self.credentials_storer(store_location="QSettings")

        # connect "Apply" button to manage
        #self.ui_auth_form.btn_ok_cancel.pressed.connect(self.manage_api_initialization)


# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == '__main__':
    """Standalone execution."""
