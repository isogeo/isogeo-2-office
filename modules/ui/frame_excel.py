# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)
# -----------------------------------------------------------------------------
# Name:         DicoGIS
# Purpose:      Automatize the creation of a dictionnary of geographic data
#               contained in a folders structures.
#               It produces an Excel output file (.xlsx)
#
# Author:       Julien Moura (@geojulien)
#
# Python:       2.7.x
# Created:      14/02/2013
# Updated:      19/03/2017
#
# Licence:      GPL 3
# ------------------------------------------------------------------------------

# ##############################################################################
# ########## Libraries #############
# ##################################

# Standard library
from Tkinter import PhotoImage, StringVar, Tk, VERTICAL
from tkinter.ttk import Entry, Label, Labelframe, Separator

import logging
from os import path

# ##############################################################################
# ############ Globals ############
# #################################

# LOG
logger = logging.getLogger("isogeo2office")

# ##############################################################################
# ########## Classes ###############
# ##################################


class FrameExcel(Labelframe):
    """Construct Excel UI."""

    def __init__(self, parent, txt=dict(), main_path="../../", ):
        """Instanciating the output workbook."""
        self.parent = parent
        Labelframe.__init__(self)

        # variables
        self.output_xl = StringVar(self)
        # self.opt_xl_join = IntVar(self)
        # self.input_xl_join_col = StringVar(self)
        # self.input_xl = ""
        # li_input_xl_cols = []

        # logo
        ico_path = path.normpath(path.join(path.abspath(main_path),
                                 "img/logo_excel2013.gif"))
        self.logo_excel = PhotoImage(master=self,
                                     file=ico_path)
        logo_excel = Label(self, borderwidth=2, image=self.logo_excel)\

        # output file
        lb_output_xl = Label(self,
                             text=_("Output filename: "))
        ent_output_xl = Entry(self,
                              textvariable=self.output_xl)

        # TO COMPLETE LATER
        # caz_xl_join = Checkbutton(fr_excel,
        #                   text=u'Joindre avec un autre fichier Excel',
        #                   variable=self.opt_xl_join,
        #                   command=lambda: self.ui_switch_xljoiner())
        # caz_xl_join.pack()

        # self.fr_input_xl_join.pack()

        # # matching with another Excel file
        # self.fr_input_xl_join = Labelframe(fr_excel,
        #                                    name='excel_joiner',
        #                                    text="Jointure à partir d'un autre tableur Excel")

        # bt_browse_input_xl = Button(self.fr_input_xl_join,
        #                             text="Choisir un fichier en entrée",
        #                             command=lambda: self.get_input_xl()).pack()
        # lb_input_xl = Label(self.fr_input_xl_join,
        #                     text=self.input_xl).pack()

        # cb_input_xl_cols = Combobox(self.fr_input_xl_join,
        #                             textvariable=self.input_xl_join_col,
        #                             values=li_input_xl_cols,
        #                             width=100)

        # griding widgets
        logo_excel.grid(row=1, rowspan=3,
                        column=0, padx=2,
                        pady=2, sticky="W")
        Separator(self, orient=VERTICAL).grid(row=1, rowspan=3,
                                              column=1, padx=2,
                                              pady=2, sticky="NSE")
        lb_output_xl.grid(row=2, column=2, sticky="W")
        ent_output_xl.grid(row=2, column=3, sticky="WE")

# #############################################################################
# ##### Stand alone program ########
# ##################################

if __name__ == '__main__':
    """To test"""
    import gettext
    root = Tk()
    frame = FrameExcel(root)
    frame.output_xl.set("isogeo2xlsx")
    frame.pack()
    root.mainloop()
