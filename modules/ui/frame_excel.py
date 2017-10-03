# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)
# -----------------------------------------------------------------------------
# Name:         Frame IsogeoToOffice - Excel
# Purpose:      Frame which is part of IsogeoToOffice UI.
#               It contains UI widgets to export in Excel.
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
from tkinter import IntVar, PhotoImage, StringVar, Tk, HORIZONTAL, VERTICAL
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


class FrameExcel(Labelframe):
    """Construct IsogeoToOffice UI Excel frame."""

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
        self.parent = parent
        Labelframe.__init__(self, text="Excel")

        # variables
        self.output_name = StringVar(self)
        self.opt_attributes = IntVar(self)
        self.opt_dashboard = IntVar(self)
        self.opt_fillfull = IntVar(self)
        self.opt_inspire = IntVar(self)

        # logo
        ico_path = path.normpath(path.join(path.abspath(main_path),
                                 "img/logo_excel2013.gif"))
        self.logo_excel = PhotoImage(master=self,
                                     file=ico_path)
        logo_excel = Label(self, borderwidth=2, image=self.logo_excel)

        # output file
        lb_output_name = Label(self,
                               text=_("Output filename: "))
        ent_output_name = Entry(self,
                                textvariable=self.output_name)

        # options
        lb_special_tabs = Label(self,
                                text=_("Goals tabs"),
                                font="Helvetica 9 bold")
        caz_dashboard = Checkbutton(self,
                                    text=_(u'Dashboard'),
                                    variable=self.opt_dashboard)
        caz_attributes = Checkbutton(self,
                                     text=_(u'Feature attributes'),
                                     variable=self.opt_attributes)
        caz_fillfull = Checkbutton(self,
                                   text=_(u'Cataloging'),
                                   variable=self.opt_fillfull)
        caz_inspire = Checkbutton(self,
                                  text=_(u'INSPIRE'),
                                  variable=self.opt_inspire)

        # griding widgets
        logo_excel.grid(row=1, rowspan=4,
                        column=0, padx=2,
                        pady=2, sticky="W")
        Separator(self, orient=VERTICAL).grid(row=1, rowspan=4,
                                              column=1, padx=2,
                                              pady=2, sticky="NSE")
        lb_output_name.grid(row=2, column=2, sticky="W")
        ent_output_name.grid(row=2, column=3, columnspan=3, sticky="WE")
        Separator(self, orient=HORIZONTAL).grid(row=3, column=2,
                                                padx=2, sticky="WE")
        lb_special_tabs.grid(row=3, column=3, columnspan=1)
        Separator(self, orient=HORIZONTAL).grid(row=3, column=4,
                                                padx=2, sticky="WE",
                                                columnspan=2)
        caz_dashboard.grid(row=4, column=2, padx=2, pady=1, sticky="W")
        caz_attributes.grid(row=4, column=3, padx=2, pady=1, sticky="E")
        caz_fillfull.grid(row=4, column=4, padx=2, pady=2)
        caz_inspire.grid(row=4, column=5, padx=2, pady=2, sticky="E")

# #############################################################################
# ##### Stand alone program ########
# ##################################

if __name__ == '__main__':
    """To test"""
    root = Tk()
    frame = FrameExcel(root)
    frame.output_name.set("isogeo2xlsx")
    frame.pack()
    root.mainloop()
