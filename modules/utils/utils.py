# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)
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

# Python 2 and 3 compatibility
from future.standard_library import install_aliases
install_aliases()

# Standard library
from collections import OrderedDict
from configparser import SafeConfigParser
from itertools import zip_longest
import logging
from os import access, path, R_OK
import re
import subprocess
from sys import platform as opersys
from time import sleep
from webbrowser import open_new_tab
from xml.sax.saxutils import escape  # '<' -> '&lt;'

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


class isogeo2office_utils(object):
    """isogeo2office utils methods class."""

    def __init__(self):
        """instanciating method
        """
        super(isogeo2office_utils, self).__init__()

        # ------------ VARIABLES ---------------------


    # MISCELLANOUS -----------------------------------------------------------

    def open_urls(self, li_url):
        """Open URLs in new tabs in the default brower.

        It waits a few seconds between the first and the next URLs
        to handle case when the webbrowser is not yet opened.
        """
        x = 1
        for url in li_url:
            if x > 1:
                sleep(3)
            else:
                pass
            open_new_tab(url)
            x += 1

        # end of method
        return

    def open_dir_file(self, target):
        """Open a file or a directory in the explorer of the operating system."""
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

    # UI --------------------------------------------------------------------

    def entry_validate_uid(self, action, index, value_if_allowed,
                           prior_value, text, validation_type,
                           trigger_type, widget_name):
        """Ensure that the users enters a boolean value in the UID option field.

        see: http://stackoverflow.com/a/8960839
        """
        if(action == '1'):
            if text in '012345678' and len(prior_value + text) < 2:
                try:
                    float(value_if_allowed)
                    return True
                except ValueError:
                    return False
            else:
                return False
        else:
            return True

    def entry_validate_date(self, action, index, value_if_allowed,
                            prior_value, text, validation_type,
                            trigger_type, widget_name):
        """Ensure that the users neters a valid value in the date option field.

        see: http://stackoverflow.com/a/8960839
        """
        if(action == '1'):
            if text in '012' and len(prior_value + text) < 2:
                try:
                    float(value_if_allowed)
                    return True
                except ValueError:
                    return False
            else:
                return False
        else:
            return True

    # SETTINGS ---------------------------------------------------------------

    def settings_load(self, config_file=r"../settings.ini"):
        """Load settings from the ini file."""
        config = SafeConfigParser()
        config.read(r"settings.ini")
        settings_dict = {s: dict(config.items(s)) for s in config.sections()}

        logger.info("Settings loaded from: {}".format(config_file))

        # end of method
        return settings_dict

    def settings_save(self, parent_ui, config_file=r"../settings.ini"):
        """Save settings into the ini file."""
        config = SafeConfigParser(dict_type=OrderedDict)
        config.read(path.realpath(config_file))
        # default OpenCatalog URL
        if len(parent_ui.shares) == 1:
            url_oc = [share[4] for share in parent_ui.shares_info[0]][0]
        else:
            url_oc = ""
            pass
        # new values to save
        config["auth"] = {"app_id": parent_ui.app_id,
                          "app_secret": parent_ui.app_secret
                          }

        config["global"] = {"out_folder": path.realpath(parent_ui.out_fold_path.get()),
                            "def_oc": url_oc,
                            "def_codelang": parent_ui.client_lang
                            }

        config["excel"] = {"excel_opt": parent_ui.opt_excel.get(),
                           "output_name": parent_ui.fr_excel.output_name.get(),
                           "opt_attributes": parent_ui.fr_excel.opt_attributes.get(),
                           "opt_fillfull": parent_ui.fr_excel.opt_fillfull.get(),
                           "opt_inspire": parent_ui.fr_excel.opt_inspire.get(),
                           }

        config["word"] = {"word_opt": parent_ui.opt_word.get(),
                          "out_prefix": parent_ui.fr_word.out_prefix.get(),
                          "tpl": parent_ui.fr_word.tpl_input.get(),
                          "opt_id": parent_ui.fr_word.opt_id.get(),
                          "opt_date": parent_ui.fr_word.opt_date.get(),
                          }

        config["xml"] = {"xml_opt": parent_ui.opt_xml.get(),
                         "opt_id": parent_ui.fr_xml.opt_id.get(),
                         "opt_date": parent_ui.fr_xml.opt_date.get(),
                         "out_prefix": parent_ui.fr_xml.out_prefix.get(),
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

    def remove_accents(self, input_str, substitute=""):
        """Clean string from special characters.

        source: http://stackoverflow.com/a/5843560
        """
        return str(substitute).join(char for char in input_str if char.isalnum())

    def clean_xml(self, invalid_xml, mode="soft", substitute="_"):
        """Clean string of XML invalid characters.

        source: http://stackoverflow.com/a/13322581/2556577
        """
        # assumptions:
        #   doc = *( start_tag / end_tag / text )
        #   start_tag = '<' name *attr [ '/' ] '>'
        #   end_tag = '<' '/' name '>'
        ws = r'[ \t\r\n]*'  # allow ws between any token
        name = '[a-zA-Z]+'  # note: expand if necessary but the stricter the better
        attr = '{name} {ws} = {ws} "[^"]*"'  # note: fragile against missing '"'; no "'"
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

    def clean_filename(self, filename, mode="soft", substitute="_"):
        """Removes invalid characters from filename."""
        if mode == "soft":
            return re.sub(r'[\\/*?:"<>|]', substitute, filename)
        elif mode == "strict":
            return re.sub("[^\w\-_\. ]", substitute, filename)
        else:
            pass

# ############################################################################
# ##### Stand alone program ########
# ##################################

if __name__ == '__main__':
    """Standalone execution and tests"""
    utils = isogeo2office_utils()

    # assert