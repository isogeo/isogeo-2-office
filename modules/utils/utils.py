# -*- coding: UTF-8 -*-
#! python3

# ----------------------------------------------------------------------------
# Name:         isogeo2office useful methods
# Purpose:      externalize util methods from isogeo2office
#
# Author:       Julien Moura (@geojulien)
#
# Python:       2.7.x
# Created:      14/08/2016
# Updated:      28/11/2016
# ----------------------------------------------------------------------------

# ############################################################################
# ########## Libraries #############
# ##################################

# Standard library
from collections import OrderedDict
from configparser import ConfigParser
from tkinter.messagebox import showerror as avert
from itertools import zip_longest
import logging
from os import access, path, R_OK
import re
import subprocess
from sys import platform as opersys
from time import sleep
from webbrowser import open_new_tab
from xml.sax.saxutils import escape  # '<' -> '&lt;'

# 3rd party
from isogeo_pysdk import IsogeoUtils

# Depending on operating system
if opersys == 'win32':
    """ windows """
    from os import startfile        # to open a folder/file
else:
    pass

# ##############################################################################
# ############ Globals ############
# #################################

logger = logging.getLogger("isogeo2office")  # LOG

# ############################################################################
# ########## Classes ###############
# ##################################


class isogeo2office_utils(IsogeoUtils):
    """isogeo2office utils methods class."""

    def __init__(self):
        """Instanciating method."""
        super(isogeo2office_utils, self).__init__()

        # ------------ VARIABLES ---------------------

    # MISCELLANOUS -----------------------------------------------------------

    def open_urls(self, li_url):
        """Open URLs in new tabs in the default brower.

        It waits a few seconds between the first and the next URLs
        to handle case when the webbrowser is not yet opened.

        :param list li_url: list of URLs to open in the default browser
        """
        x = 1
        for url in li_url:
            if x > 1:
                sleep(3)
            else:
                pass
            open_new_tab(url)
            x += 1

    def open_dir_file(self, target):
        """Open a file or a directory in the explorer of the operating system.

        :param str target: path of the folder or file to open
        """
        # check if the file or the directory exists
        if not path.exists(target):
            raise IOError('No such file: {0}'.format(target))

        # check the read permission
        if not access(target, R_OK):
            raise IOError('Cannot access file: {0}'.format(target))

        # open the directory or the file according to the os
        if opersys == 'win32':  # Windows
            proc = startfile(path.realpath(target))

        elif opersys.startswith('linux'):  # Linux:
            proc = subprocess.Popen(['xdg-open', target],
                                    stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE)

        elif opersys == 'darwin':  # Mac:
            proc = subprocess.Popen(['open', '--', target],
                                    stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE)

        else:
            raise NotImplementedError(
                "Your `%s` isn't a supported operating system`." % opersys)

        # end of function
        return proc

    # ISOGEO -----------------------------------------------------------------

    def get_url_base(self, url_input):
        """Get OpenCatalog base URL to add resource ID easily."""
        # get the OpenCatalog URL given
        if not url_input[-1] == '/':
            url_input = url_input + '/'
        else:
            pass

        # get the clean url
        return url_input[0:url_input.index(url_input.rsplit('/')[6])]

    # SETTINGS ---------------------------------------------------------------

    def settings_save(self, parent, config_file=r"../settings.ini"):
        """Save settings into the ini file."""
        config = ConfigParser(dict_type=OrderedDict)
        config.read(path.realpath(config_file))
        # default OpenCatalog URL
        if len(parent.shares) == 1:
            url_oc = [share[4] for share in parent.shares_info[0]][0]
        else:
            url_oc = ""
            pass
        # new values to save
        config["auth"] = {"app_id": parent.app_id,
                          "app_secret": parent.app_secret
                          }

        config["local"] = {"out_folder": path.realpath(parent.out_fold_path.get()),
                           "def_oc": url_oc,
                           "def_codelang": parent.client_lang
                           }

        config["excel"] = {"excel_opt": parent.opt_excel.get(),
                           "output_name": parent.fr_excel.output_name.get(),
                           "opt_dashboard": parent.fr_excel.opt_dashboard.get(),
                           "opt_attributes": parent.fr_excel.opt_attributes.get(),
                           "opt_fillfull": parent.fr_excel.opt_fillfull.get(),
                           "opt_inspire": parent.fr_excel.opt_inspire.get(),
                           }

        config["word"] = {"word_opt": parent.opt_word.get(),
                          "out_prefix": parent.fr_word.out_prefix.get(),
                          "tpl": parent.fr_word.tpl_input.get(),
                          "opt_id": parent.fr_word.opt_id.get(),
                          "opt_date": parent.fr_word.opt_date.get(),
                          }

        config["xml"] = {"xml_opt": parent.opt_xml.get(),
                         "out_prefix": parent.fr_xml.out_prefix.get(),
                         "opt_id": parent.fr_xml.opt_id.get(),
                         "opt_date": parent.fr_xml.opt_date.get(),
                         "opt_zip": parent.fr_xml.opt_zip.get(),
                         }

        # writing
        with open(path.realpath(config_file), mode="w") as configfile:
            try:
                config.write(configfile)
                logger.info("Settings saved into: {}".format(config_file))
            except UnicodeEncodeError as e:
                avert(_("Invalid character"),
                      _("Special character spotted in output filenames.\n"
                      "Settings couldn't be saved but exports will continue."))
                logger.error("SETTINGS - Encoding error: {}".format(e))

        # end of method
        return

    # ------------------------------------------------------------------------
    def clean_filename(self, filename: str, substitute: str = "", mode: str = "soft"):
        """Remove invalid characters from filename.
        \\ TO DO: isnt' duplicated with next method on special chars?

        :param str filename: filename string to clean
        :param str substitute: character to use for subtistution of special chars
        :param str modeaccents: mode to apply. Available options:

          * soft [default]: remove chars which are not accepted in filenames
          * strict: remove additional chars (punctuation...)
        """
        if mode == "soft":
            return re.sub(r'[\\/*?:"<>|]', substitute, filename)
        elif mode == "strict":
            return re.sub(r"[^\w\-_\. ]", substitute, filename)
        else:
            raise ValueError("'mode' option must be one of: soft | strict")

    def clean_special_chars(self, input_str: str, substitute: str = "", accents: bool = 1):
        """Clean string from special characters.

        Source: https://stackoverflow.com/a/38799620/2556577

        :param str input_str: string to clean
        :param str substitute: character to use for subtistution of special chars
        :param bool accents: option to keep or not the accents
        """
        if accents:
            return re.sub(r'\W+', substitute, input_str)
        else:
            return re.sub(r'[^A-Za-z0-9]+', substitute, input_str)

    def clean_xml(self, invalid_xml, mode: str = "soft", substitute: str = "_"):
        """Clean string of XML invalid characters.

        source: https://stackoverflow.com/a/13322581/2556577

        :param str invalid_xml: xml string to clean
        :param str substitute: character to use for subtistution of special chars
        :param str modeaccents: mode to apply. Available options:

          * soft [default]: remove chars which are not accepted in XML
          * strict: remove additional chars
        """
        # assumptions:
        #   doc = *( start_tag / end_tag / text )
        #   start_tag = '<' name *attr [ '/' ] '>'
        #   end_tag = '<' '/' name '>'
        ws = r'[ \t\r\n]*'  # allow ws between any token
        # note: expand if necessary but the stricter the better
        name = '[a-zA-Z]+'
        # note: fragile against missing '"'; no "'"
        attr = '{name} {ws} = {ws} "[^"]*"'
        start_tag = '< {ws} {name} {ws} (?:{attr} {ws})* /? {ws} >'
        end_tag = '{ws}'.join(['<', '/', '{name}', '>'])
        tag = '{start_tag} | {end_tag}'

        assert '{{' not in tag
        while '{' in tag:   # unwrap definitions
            tag = tag.format(**vars())

        tag_regex = re.compile('(%s)' % tag, flags=re.VERBOSE)

        # escape &, <, > in the text
        iters = [iter(tag_regex.split(invalid_xml))] * 2
        pairs = zip_longest(*iters, fillvalue='')  # iterate 2 items at a time

        # get the clean version
        clean_version = ''.join(escape(text) + tag for text, tag in pairs)
        if mode == "strict":
            clean_version = re.sub(r"<.*?>", substitute, clean_version)
        else:
            pass
        return clean_version


# ############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == '__main__':
    """Standalone execution and tests"""
    utils = isogeo2office_utils()
