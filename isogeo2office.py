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
from Tkinter import Tk, StringVar, IntVar    # GUI
from ttk import Label, Button, Entry, Combobox, Labelframe   # widgets

# 3rd party library


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
    def __init__(self, _lang):
        super(Isogeo2office, self).__init__()
        # ------------ Settings ----------------
        self.settings_load()
        app_id = self.setting.get('auth').get('app_id')
        app_secret = self.setting.get('auth').get('app_secret')
        client_lang = self.setting.get('basics').get('def_codelang')

        # ------------ Isogeo authentification ----------------
        self.isogeo = Isogeo(client_id=app_id,
                             client_secret=app_secret,
                             lang=client_lang)
        self.token = self.isogeo.connect()

        # ------------ Variables ----------------
        li_tpls = [path.abspath(path.join(r'templates', tpl)) for tpl in listdir(r'templates') if path.splitext(tpl)[1].lower() == ".docx"]

        # ------------ UI ----------------
        self.title('isogeo2office - ToolBox')

        # Frames
        self.FrGlobal = Labelframe(self,
                                   name='global',
                                   text='Générique')

        self.FrExcel = Labelframe(self,
                                  name='excel',
                                  text='Fichier Excel')

        self.FrWord = Labelframe(self,
                                 name='word',
                                 text='Fichier Word')

        # ## GLOBAL ##
        url_input = StringVar(self)

        # étiquette
        lb_input_oc = Label(self.FrGlobal,
                            text="Coller l'URL d'un OpenCatalog").pack()
        ent_OpenCatalog = Entry(FrGlobal,
                                textvariable=url_input,
                                width=100)
        # ent_OpenCatalog.insert(0, "https://open.isogeo.com/s/ad6451f1f9ca405ca6f78fabf46aeb10/Bue0ySfhmGOPw33jHMyaJtcOM4MY0/q/keyword:inspire-theme:administrativeunits")
        ent_OpenCatalog.pack()
        ent_OpenCatalog.focus_set()

        droplist.pack()

        # ## EXCEL ##


        # ## WORD ##
        # variables
        tpl_input = StringVar(self)
        # pick a template
        lb_input_tpl = Label(self.FrWord,
                             text="Choisir un template").pack()
        droplist = Combobox(self.FrWord,
                            textvariable=tpl_input,
                            values=li_tpls,
                            width=100)
        Button(self.FrWord,
               text="Wordification !",
               command=lambda: app.destroy()).pack()


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
