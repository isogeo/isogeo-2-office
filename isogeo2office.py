# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)
#------------------------------------------------------------------------------
# Name:         Isogeo
# Purpose:      Get metadatas from an Isogeo share and store it into files
#
# Author:       Julien Moura (@geojulien)
#
# Python:       2.7.x
# Created:      18/12/2015
# Updated:      22/01/2016
#------------------------------------------------------------------------------

###############################################################################
########### Libraries #############
###################################

# Standard library
import ConfigParser
from datetime import datetime
from os import listdir, path
from sys import exit
from tkFileDialog import askopenfilename
from Tkinter import Tk, StringVar, IntVar, Frame    # GUI
from ttk import Label, Button, Entry, Combobox, Labelframe, Checkbutton   # widgets

# 3rd party library
from openpyxl import load_workbook

# Custom modules
from modules.isogeo_sdk import Isogeo
from modules.isogeo2xls import Isogeo2xlsx
from modules.isogeo2docx import Isogeo2docx

###############################################################################
########### Classes ###############
###################################


class Isogeo2office(Tk):
    """
    docstring for Isogeo to Office
    """
    def __init__(self):
        Tk.__init__(self)
        # ------------ Settings ----------------
        self.settings_load()
        app_id = self.settings.get('auth').get('app_id')
        app_secret = self.settings.get('auth').get('app_secret')
        client_lang = self.settings.get('basics').get('def_codelang')

        # ------------ Isogeo authentification ----------------
        self.isogeo = Isogeo(client_id=app_id,
                             client_secret=app_secret,
                             lang=client_lang)
        self.token = self.isogeo.connect()

        # # ------------ Isogeo search ----------------
        # self.search_results = self.isogeo.search(self.token,
        #                                          sub_resources=self.isogeo.sub_resources_available,
        #                                          preprocess=1)

        # ------------ Variables ----------------
        li_tpls = [path.abspath(path.join(r'templates', tpl)) for tpl in listdir(r'templates') if path.splitext(tpl)[1].lower() == ".docx"]

        # ------------ UI ----------------
        self.title('isogeo2office - ToolBox')

        # Frames
        fr_global = Labelframe(self,
                               name='global',
                               text="Générique")

        fr_excel = Labelframe(self,
                              name='excel',
                              text="Export Excel")

        fr_word = Labelframe(self,
                             name='word',
                             text="Export Word")

        # ## GLOBAL ##
        url_input = StringVar(self)

        # lb_count_avail_resources = Label(fr_global,
        #                                  text="{} métadonnées partagées".format(self.search_results.get('total'))).pack()

        # OpenCatalog URL
        lb_input_oc = Label(fr_global,
                            text="Coller l'URL d'un OpenCatalog").pack()
        ent_opencatalog = Entry(fr_global,
                                textvariable=url_input,
                                width=100)
        # ent_OpenCatalog.insert(0, "https://open.isogeo.com/s/ad6451f1f9ca405ca6f78fabf46aeb10/Bue0ySfhmGOPw33jHMyaJtcOM4MY0/q/keyword:inspire-theme:administrativeunits")
        ent_opencatalog.pack()
        ent_opencatalog.focus_set()

        fr_global.pack()

        # ------------------------------------------------------------

        # ## EXCEL ##
        # variables
        output_xl = StringVar(self)
        self.opt_xl_join = IntVar(fr_excel)
        self.input_xl = ""
        li_input_xl_cols = []
        self.input_xl_join_col = StringVar()

        # output file
        lb_output_xl = Label(fr_excel,
                             text="Nom du fichier en sortie: ").pack()
        ent_output_xl = Entry(fr_excel,
                              text="Nom du fichier en sortie: ",
                              textvariable=output_xl,
                              width=100).pack()

        # matching with another Excel file
        self.fr_input_xl_join = Labelframe(fr_excel,
                                           name='excel_joiner',
                                           text="Jointure à partir d'un autre tableur Excel")
        caz_xl_join = Checkbutton(fr_excel,
                                  text=u'Joindre avec un autre fichier Excel',
                                  variable=self.opt_xl_join,
                                  command=lambda: self.ui_switch_xljoiner())
        caz_xl_join.pack()

        bt_browse_input_xl = Button(self.fr_input_xl_join,
                                    text="Choisir un fichier en entrée",
                                    command=lambda: self.get_input_xl()).pack()
        lb_input_xl = Label(self.fr_input_xl_join,
                            text=self.input_xl).pack()

        cb_input_xl_cols = Combobox(self.fr_input_xl_join,
                                    textvariable=self.input_xl_join_col,
                                    values=li_input_xl_cols,
                                    width=100)
        cb_input_xl_cols.pack()


        self.fr_input_xl_join.pack()

        Button(fr_excel,
               text="Excelization !",
               command=lambda: process_excelization()).pack()

        fr_excel.pack()

        # ------------------------------------------------------------

        # ## WORD ##
        # variables
        tpl_input = StringVar(self)
        # pick a template
        lb_input_tpl = Label(fr_word,
                             text="Choisir un template").pack()
        cb_available_tpl = Combobox(fr_word,
                                    textvariable=tpl_input,
                                    values=li_tpls,
                                    width=100)
        cb_available_tpl.pack()

        Button(fr_word,
               text="Wordification !",
               command=lambda: process_wordification()).pack()
        # packing frame
        fr_word.pack()
        # ------------------------------------------------------------

    def get_input_xl(self):
        """
        """
        self.input_xl = askopenfilename(parent=self,
                                        filetypes=[("Excel 2010 files","*.xlsx"),("Excel 2003 files","*.xls")],
                                        title=u"Choisir le fichier Excel à partir duquel faire la jointure")

        if self.input_xl:
            print(self.input_xl)
            return
        else:
            print(u'Aucun fichier sélectionné')

    def ui_switch_xljoiner(self):
        """ Enable/disable the form for input xl to join.
        """
        if self.opt_xl_join.get():
            self.fr_input_xl_join.pack()
        else:
            self.fr_input_xl_join.pack_forget()
        # end of function
        return


    def get_basic_metrics(self):
        """ TO DO
        """
        empty_search = self.isogeo.search(self.token,
                                          # query="keyword:isogeo:2015",
                                          page_size=0,
                                          whole_share=0,
                                          prot='http')

        # end of method
        return len(empty_search.get('results'))


    def get_search_results(self):
        """ TO DO
        """
        pass

    def settings_load(self):
        """ TO DO
        """
        config = ConfigParser.SafeConfigParser()
        config.read(r"settings.ini")
        self.settings = {s:dict(config.items(s)) for s in config.sections()}

        # end of method
        return

    def settings_save(self, ):
        """ TO DO
        """

        # end of method
        return

    def get_url_base(self, url_input):
        """ TO DO
        """
        # get the OpenCatalog URL given
        if not url_input[-1] == '/':
            url_input = url_input + '/'
        else:
            pass

        # get the clean url
        url_output = url_input[0:url_input.index(url_input.rsplit('/')[6])]

        # end of method
        return url_output

    def process_excelization(self, id_resource):
        """ TO DO
        """

        # Sauvegarde du fichier Excel
        dstamp = datetime.now()
        book.save(r"output\isogeo2xls_{0}_{1}{2}{3}{4}{5}{6}.xls".format(share_rez.get("name"),
                                                                         dstamp.year,
                                                                         dstamp.month,
                                                                         dstamp.day,
                                                                         dstamp.hour,
                                                                         dstamp.minute,
                                                                         dstamp.second))

        # end of method
        return

    def process_wordification(self, search_results):
        """ TO DO
        """
        ## WORDIZING METADATAS #################
        for md in metadatas:
            docx_tpl = DocxTemplate(path.realpath(tpl_input.get()))
            dstamp = datetime.now()
            md2docx(docx_tpl, 0, md, li_catalogs, url_base)  # passing parameters to the Word generator
            docx_tpl.save(r"output\{0}_{8}_{7}_{1}{2}{3}{4}{5}{6}.docx".format(share_rez.get("name"),
                                                                           dstamp.year,
                                                                           dstamp.month,
                                                                           dstamp.day,
                                                                           dstamp.hour,
                                                                           dstamp.minute,
                                                                           dstamp.second,
                                                                           md.get("_id")[:5],
                                                                           remove_accents(md.get("title")[:15], "_")))


        # end of method
        return

# ###############################################################################
# ###### Stand alone program ########
# ###################################

if __name__ == '__main__':
    """ standalone execution """
    app = Isogeo2office()
    app.mainloop()
