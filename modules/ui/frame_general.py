# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)
# -----------------------------------------------------------------------------
# Name:         Frame IsogeoToOffice - Buttons
# Purpose:      Frame which is part of IsogeoToOffice UI.
#               It contains buttons to open links in a web browser.
# Author:       Julien Moura (@geojulien)
#
# Python:       2.7.x
# Created:      14/02/2015
# Updated:      19/10/2017
#
# Licence:      GPL 3
# ------------------------------------------------------------------------------

# ##############################################################################
# ########## Libraries #############
# ##################################

# Standard library
import gettext  # localization
import logging
from tkinter import Tk
from tkinter.ttk import Button, Frame
from webbrowser import open_new_tab

# ##############################################################################
# ############ Globals ############
# #################################

logger = logging.getLogger("isogeo2office")  # LOG

# ##############################################################################
# ########## Classes ###############
# ##################################


class FrameGlobal(Frame):
    """Construct IsogeoToOffice UI global frame."""

    def __init__(self, parent, main_path="../../", lang=None):
        """Instanciating the output workbook."""
        # localization
        try:
            # lang.install(unicode=1)
            _ = lang.gettext
            logger.info("Custom language set: {}"
                        .format(_("English")))
        except Exception as e:
            logger.error(e)
            _ = gettext.gettext
            logger.info("Default language set: English")

        # UI
        self.master = parent
        self.parent = parent
        logger.debug(parent)
        Frame.__init__(self)

        # WIDGETS
        # for unicode symbols: https://www.w3schools.com/charsets/ref_utf_symbols.asp
        # help
        url_help = "https://isogeo.gitbooks.io/app-isogeo2office/content/"
        self.btn_help = Button(self,
                               text="\U00002753 {}".format(_("Help")),
                               command=lambda: open_new_tab(url_help))
        # contact
        mailto = _("mailto:Isogeo%20Projects%20"
                   "<projects+isogeo2office@isogeo.com>?"
                   "subject=[Isogeo2office]%20Question")
        self.btn_contact = Button(self,
                                  text="\U00002709 {}".format(_("Contact")),
                                  command=lambda: open_new_tab(mailto))
        # authentication
        self.btn_settings = Button(self,
                                   text="\U000026BF {}".format(_("Settings")),
                                   command=lambda: parent.ui_settings_prompt(lang))
        # shares
        self.btn_open_shares = Button(self,
                                      text="\U00002692 {}".format(_("Admin shares")),
                                      command=lambda: self.utils.open_urls(li_oc))
        # source
        url_src = "https://github.com/isogeo/isogeo-2-office/issues/new"
        self.btn_src = Button(self,
                              text="\U000026A0 {}".format(_("Report")),
                              command=lambda: open_new_tab(url_src))

        # griding widgets
        self.btn_help.grid(row=1,
                           column=0, padx=2, pady=2,
                           sticky="NWE")
        self.btn_contact.grid(row=1,
                              column=1, padx=2, pady=2,
                              sticky="NWE")
        self.btn_open_shares.grid(row=1,
                                  column=2, padx=2, pady=2,
                                  sticky="NWE")
        self.btn_settings.grid(row=1,
                               column=3, padx=2, pady=2,
                               sticky="NWE")
        self.btn_src.grid(row=1,
                          column=4, padx=2, pady=2,
                          sticky="NWE")


# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == '__main__':
    """To test"""
    root = Tk()
    frame = FrameGlobal(root)
    frame.pack()
    root.mainloop()
