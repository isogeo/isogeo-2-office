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
import gettext  # localization
from tkinter import PhotoImage, IntVar, StringVar, Tk, VERTICAL
from tkinter.ttk import Entry, Label, Labelframe, Separator, Combobox

import logging
from os import listdir, path

# ##############################################################################
# ############ Globals ############
# #################################

logger = logging.getLogger("isogeo2office")  # LOG

# ##############################################################################
# ########## Classes ###############
# ##################################


class FrameWord(Labelframe):
    """Construct Excel UI."""

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
        Labelframe.__init__(self, text="Word")

        # variables
        li_tpls = [tpl for tpl in listdir(path.join(path.abspath(main_path), r'templates'))
                   if path.splitext(tpl)[1].lower() == ".docx"]
        self.tpl_input = StringVar(self)
        self.out_prefix = StringVar(self, "isogeo2docx")
        self.opt_id = IntVar(self, 5)
        self.opt_date = IntVar(self, 1)
        # fields validation
        val_uid = validators.get("val_uid")
        val_date = validators.get("val_date")

        # logo
        ico_path = path.normpath(path.join(path.abspath(main_path),
                                 "img/logo_word2013.gif"))
        self.logo_word = PhotoImage(master=self,
                                     file=ico_path)
        logo_word = Label(self, borderwidth=2, image=self.logo_word)

        # pick a template
        lb_input_tpl = Label(self,
                             text=_("Pick a template: "))
        cb_available_tpl = Combobox(self,
                                    textvariable=self.tpl_input,
                                    values=li_tpls)

        prev_tpl = ""
        # prev_tpl = self.settings.get("word").get("word_tpl")
        cb_available_tpl.current(li_tpls.index(prev_tpl)\
                                 if prev_tpl in li_tpls else 0)
        # specific options
        lb_out_word_prefix = Label(self, text=_("File prefix: "))
        lb_out_word_uid = Label(self, text=_("UID chars:\n"
                                             "(0 - 8)"))
        lb_out_word_date = Label(self, text=_("Timestamp:\n"
                                              "(0=no, 1=date, 2=datetime)"))

        ent_out_word_prefix = Entry(self, textvariable=self.out_prefix)
        ent_out_word_uid = Entry(self, textvariable=self.opt_id,
                                 width=2, validate="key",
                                 validatecommand=val_uid)
        ent_out_word_date = Entry(self, textvariable=self.opt_date,
                                  width=2, validate="key",
                                  validatecommand=val_date)

        # griding widgets
        logo_word.grid(row=1, rowspan=3,
                       column=0, padx=2,
                       pady=2, sticky="W")
        Separator(self, orient=VERTICAL).grid(row=1, rowspan=3,
                                              column=1, padx=2,
                                              pady=2, sticky="NSE")
        lb_input_tpl.grid(row=1, column=2, padx=2, pady=2, sticky="W")
        cb_available_tpl.grid(row=1, column=3, columnspan=2,
                              padx=2, pady=2, sticky="WE")
        lb_out_word_prefix.grid(row=2, column=2, padx=2, pady=2, sticky="W")
        ent_out_word_prefix.grid(row=2, column=3, columnspan=2,
                                 padx=2, pady=2, sticky="WE")

        lb_out_word_uid.grid(row=3, column=2, padx=2, pady=2, sticky="W")
        ent_out_word_uid.grid(row=3, column=2, padx=3, pady=2, sticky="E")
        lb_out_word_date.grid(row=3, column=3, padx=3, pady=2, sticky="W")
        ent_out_word_date.grid(row=3, column=4, padx=2, pady=2, sticky="W")

# #############################################################################
# ##### Stand alone program ########
# ##################################

if __name__ == '__main__':
    """To test"""
    root = Tk()
    frame = FrameWord(root, validators=dict())
    frame.pack()
    root.mainloop()
