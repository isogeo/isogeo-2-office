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

# Standard library
from ConfigParser import SafeConfigParser
import logging
from webbrowser import open_new_tab

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

        logging.info("Settings loaded from: {}".format(config_file))

        # end of method
        return settings_dict

    def settings_save(self, config_file=r"../settings.ini"):
        """Save settings into the ini file."""
        config = SafeConfigParser()
        config.read(path.realpath(config_file))
        # new values
        config.set('auth', 'app_id', self.app_id)
        config.set('auth', 'app_secret', self.app_secret)
        config.set('basics', 'out_folder', path.realpath(self.out_fold_path.get()))
        config.set('basics', 'excel_out', self.output_xl.get())
        config.set('basics', 'excel_opt', str(self.opt_excel.get()))
        config.set('basics', 'word_opt', str(self.opt_word.get()))
        config.set('basics', 'word_tpl', self.tpl_input.get())
        config.set('basics', 'word_out_prefix', str(self.out_word_prefix.get()))
        config.set('basics', 'word_opt_id', str(self.word_opt_id.get()))
        config.set('basics', 'word_opt_date', str(self.word_opt_date.get()))
        config.set('basics', 'xml_opt', str(self.opt_xml.get()))
        config.set('basics', 'xml_out_prefix', str(self.out_xml_prefix.get()))
        config.set('basics', 'xml_opt_id', str(self.xml_opt_id.get()))
        config.set('basics', 'xml_opt_date', str(self.xml_opt_date.get()))
        # writing
        with open(path.realpath(config_file), 'wb') as configfile:
            config.write(configfile)

        logging.info("Settings saved into: {}".format(config_file))
        # end of method
        return

    # ------------------------------------------------------------------------



# ############################################################################
# ##### Stand alone program ########
# ##################################

if __name__ == '__main__':
    """Standalone execution and tests"""
    utils = isogeo2office_utils()

    # assert