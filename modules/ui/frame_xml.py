# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)
# -----------------------------------------------------------------------------
# Name:         Frame IsogeoToOffice - XML export
# Purpose:      Frame which is part of IsogeoToOffice UI.
#               It contains UI widgets to export into XML.
#
# Author:       Julien Moura (@geojulien)
#
# Python:       2.7.x
# Created:      14/02/2015
# Updated:      19/03/2017
#
# Licence:      GPL 3
# ------------------------------------------------------------------------------

# ##############################################################################
# ########## Libraries #############
# ##################################

# Standard library
import gettext  # localization
from tkinter import IntVar, PhotoImage, StringVar, Tk, VERTICAL
from tkinter.ttk import Entry, Label, Labelframe, Separator, Checkbutton

import logging
from os import path

# ##############################################################################
# ############ Globals ############
# #################################

logger = logging.getLogger("isogeo2office")  # LOG

# ##############################################################################
# ########## Classes ###############
# ##################################


class FrameXml(Labelframe):
    """Construct IsogeoToOffice UI XML frame."""

    def __init__(self, parent, main_path="../../", lang=None, validators=None):
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
        self.parent = parent
        Labelframe.__init__(self, text="XML")

        # variables
        self.out_prefix = StringVar(self)
        self.opt_id = IntVar(self)
        self.opt_date = IntVar(self)
        self.opt_zip = IntVar(self)
        # fields validation
        val_uid = validators.get("val_uid")
        val_date = validators.get("val_date")

        # logo
        ico_path = path.normpath(path.join(path.abspath(main_path),
                                 "img/logo_inspireFun.gif"))
        self.logo_xml = PhotoImage(master=self,
                                   file=ico_path)
        logo_xml = Label(self, borderwidth=2, image=self.logo_xml)

        # options
        caz_zip_xml = Checkbutton(self,
                                  text=_(u'Pack all XML into one ZIP'),
                                  variable=self.opt_zip)
        lb_out_xml_prefix = Label(self, text=_("File prefix: "))
        lb_out_xml_uid = Label(self, text=_("UID chars:\n"
                                            "(0 - 8)"))
        lb_out_xml_date = Label(self, text=_("Timestamp:\n"
                                             "(0=no, 1=date, 2=datetime)"))

        ent_out_xml_prefix = Entry(self, textvariable=self.out_prefix)
        ent_out_xml_uid = Entry(self, textvariable=self.opt_id,
                                width=2, validate="key",
                                validatecommand=val_uid)
        ent_out_xml_date = Entry(self, textvariable=self.opt_date,
                                 width=2, validate="key",
                                 validatecommand=val_date)

        # griding widgets
        logo_xml.grid(row=1, rowspan=3,
                      column=0, padx=2,
                      pady=2, sticky="W")
        Separator(self, orient=VERTICAL).grid(row=1, rowspan=3,
                                              column=1, padx=2,
                                              pady=2, sticky="NSE")
        lb_out_xml_prefix.grid(row=1, column=2, padx=2, pady=2, sticky="W")
        ent_out_xml_prefix.grid(row=1, column=3, columnspan=2,
                                padx=2, pady=2, sticky="WE")

        lb_out_xml_uid.grid(row=2, column=2, padx=2, pady=2, sticky="W")
        ent_out_xml_uid.grid(row=2, column=3, padx=0, pady=2, sticky="E")
        lb_out_xml_date.grid(row=2, column=4, padx=3, pady=2, sticky="W")
        ent_out_xml_date.grid(row=2, column=5, padx=2, pady=2, sticky="E")
        caz_zip_xml.grid(row=3, column=2, columnspan=3, padx=2, pady=2, sticky="WE")


# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == '__main__':
    """To test"""
    root = Tk()
    frame = FrameXml(root, validators=dict())
    frame.pack()
    root.mainloop()
