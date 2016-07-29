# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)
# ----------------------------------------------------------------------------
# Name:         Isogeo
# Purpose:      Get metadatas from an Isogeo share and store it into files
#
# Author:       Julien Moura (@geojulien)
#
# Python:       2.7.x
# Created:      18/12/2015
# Updated:      22/01/2016
# ---------------------------------------------------------------------------

# ############################################################################
# ########## Libraries #############
# ##################################

# Standard library
from ConfigParser import SafeConfigParser
from datetime import datetime
import logging      # log files
from logging.handlers import RotatingFileHandler
from os import listdir, path
from sys import argv, exit
from time import sleep
from tkFileDialog import askopenfilename
from Tkinter import Tk, StringVar, IntVar, Image, PhotoImage   # GUI
from ttk import Label, Button, Entry, Checkbutton, Combobox  # advanced widgets
from ttk import Labelframe, Progressbar, Style  # advanced widgets
from webbrowser import open_new_tab

# 3rd party library
from isogeo_pysdk import Isogeo
from openpyxl import load_workbook
import requests

# Custom modules
from modules.isogeo2xlsx import Isogeo2xlsx
from modules.isogeo2docx import Isogeo2docx
from modules import CheckNorris

# ############################################################################
# ########## Global ################
# ##################################

# VERSION
_version = "1.1"

# LOG FILE ##
# see: http://sametmax.com/ecrire-des-logs-en-python/
logger = logging.getLogger()
logging.captureWarnings(True)
logger.setLevel(logging.DEBUG)  # all errors will be get
log_form = logging.Formatter('%(asctime)s || %(levelname)s || %(module)s || %(message)s')
logfile = RotatingFileHandler('isogeo2office.log', 'a', 5000000, 1)
logfile.setLevel(logging.DEBUG)
logfile.setFormatter(log_form)
logger.addHandler(logfile)
logger.info('\t============== Isogeo => Office =============')


# ############################################################################
# ########## Classes ###############
# ##################################

class Isogeo2office(Tk):
    """ UI Class to
    docstring for Isogeo to Office
    """
    # attributes and global actions
    logger.info('Version: {0}'.format(_version))

    def __init__(self, ui_launcher=1):
        """
        """
        # Invoke Check Norris
        checker = CheckNorris()

        # checking connection
        if not checker.check_internet_connection():
            logger.error('An Internet connection is required. Check your settings.')
            exit()
        else:
            pass

        # UI or not to UI
        if not ui_launcher:
            self.no_ui_launcher()
        else:
            pass

        Tk.__init__(self)

        # ------------ Settings ----------------
        self.settings_load()
        app_id = self.settings.get('auth').get('app_id')
        app_secret = self.settings.get('auth').get('app_secret')
        client_lang = self.settings.get('basics').get('def_codelang')
        def_oc = self.settings.get('basics').get('def_oc')

        # ------------ Isogeo authentification ----------------
        self.isogeo = Isogeo(client_id=app_id,
                             client_secret=app_secret,
                             lang=client_lang)
        self.token = self.isogeo.connect()

        # ------------ Isogeo search & shares ----------------
        self.search_results = self.isogeo.search(self.token,
                                                 page_size=0,
                                                 whole_share=0)
        self.shares = self.isogeo.shares(self.token)
        self.shares_info = self.get_shares_info()

        # ------------ Variables ----------------
        li_tpls = [path.abspath(path.join(r'templates', tpl))
                   for tpl in listdir(r'templates')
                   if path.splitext(tpl)[1].lower() == ".docx"]

        # ------------ UI ----------------
        self.title("isogeo2office - {0}".format(_version))
        icon = Image("photo", file=r"img/favicon_isogeo.gif")
        self.call("wm", "iconphoto", self._w, icon)
        self.style = Style().theme_use("clam")
        self.resizable(width=False, height=False)
        self.focus_force()

        # Frames
        fr_isogeo = Labelframe(self, name="isogeo", text="Isogeo")
        fr_excel = Labelframe(self, name="excel", text="Excel")
        fr_word = Labelframe(self, name="word", text="Word")
        fr_process = Labelframe(self, name="process", text="Lancer")

        fr_isogeo.grid(row=1, column=1, sticky="WE")
        fr_excel.grid(row=2, column=1, sticky="WE")
        fr_word.grid(row=3, column=1, sticky="WE")
        fr_process.grid(row=4, column=1, sticky="WE")

        
        # ------------------------------------------------------------

        # ## GLOBAL ##
        self.app_metrics = StringVar(fr_isogeo)
        self.oc_msg = StringVar(fr_isogeo)
        self.url_input = StringVar(fr_isogeo)

        # logo
        self.logo_isogeo = PhotoImage(master=fr_isogeo, file=r'img/logo_isogeo.gif')
        logo_isogeo = Label(fr_isogeo, borderwidth=2, image=self.logo_isogeo)

        # metrics
        self.app_metrics.set("{} métadonnées partagées via {} partages,\nappartenant à {} groupes de travail différents."\
                             .format(self.search_results.get('total'),
                                     len(self.shares),
                                     len(self.shares_info[1])))
        lb_app_metrics = Label(fr_isogeo,
                               textvariable=self.app_metrics)

        # OpenCatalog to display
        self.lb_input_oc = Label(fr_isogeo,
                                 textvariable=self.oc_msg)
        ent_opencatalog = Entry(fr_isogeo,
                                textvariable=self.url_input,
                                width=100)

        if len(self.shares_info[2]) != 0:
            logger.info("Any OpenCatalog found among the shares")
            self.oc_msg.set("{} partages à cette application n'ont pas d'OpenCatalog."
                            "\nAjouter l'application OpenCatalog en suivant les liens ci-dessous."
                            "\nPuis redémarrer l'application.".format(len(self.shares_info[2])))
            btn_open_shares = Button(fr_isogeo,
                                     text="Corriger les partages",
                                     command=lambda: self.open_urls(self.shares_info[2]))            
        else:
            logger.info("All shares have an OpenCatalog")
            self.oc_msg.set("Configuration OK.")
            li_oc = [share[3] for share in self.shares_info[0]]
            btn_open_shares = Button(fr_isogeo,
                                     text="Consulter les partages",
                                     command=lambda: self.open_urls(li_oc))

        # griding widgets
        logo_isogeo.grid(row=1, rowspan=3,
                    column=0, padx=2,
                    pady=2, sticky="W")
        lb_app_metrics.grid(row=1, column=1, sticky="WE")
        self.lb_input_oc.grid(row=2, column=1, sticky="WE")
        btn_open_shares.grid(row=2, column=2, sticky="WE")

        # ------------------------------------------------------------

        # ## EXCEL ##
        # variables
        output_xl = StringVar(self)
        self.opt_xl_join = IntVar(fr_excel)
        self.input_xl_join_col = StringVar(fr_excel)
        self.input_xl = ""
        li_input_xl_cols = []

        # logo
        self.logo_excel = PhotoImage(master=fr_excel, file=r'img/logo_excel2013.gif')
        logo_excel = Label(fr_excel, borderwidth=2, image=self.logo_excel)\

        # output file
        lb_output_xl = Label(fr_excel,
                             text="Nom du fichier en sortie: ")
        ent_output_xl = Entry(fr_excel,
                              text="Nom du fichier en sortie: ",
                              textvariable=output_xl)

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
        lb_output_xl.grid(row=1, column=1)
        ent_output_xl.grid(row=2, column=1)

        # ------------------------------------------------------------

        # ## WORD ##
        # variables
        self.tpl_input = StringVar(self)
        
        # logo
        self.logo_word = PhotoImage(master=fr_word, file=r'img/logo_word2013.gif')
        logo_word = Label(fr_word, borderwidth=2, image=self.logo_word)
        
        # pick a template
        lb_input_tpl = Label(fr_word,
                             text="Choisir un template")
        cb_available_tpl = Combobox(fr_word,
                                    textvariable=self.tpl_input,
                                    values=li_tpls)
      
        # griding widgets
        logo_word.grid(row=1, rowspan=3,
                       column=0, padx=2,
                       pady=2, sticky="W")
        lb_input_tpl.grid(row=1, column=1)
        cb_available_tpl.grid(row=2, column=1)

        # ------------------------------------------------------------

        # ## PROCESS ##
        # variables
        self.opt_excel = IntVar(fr_process)
        self.opt_word = IntVar(fr_process)

        # logo
        self.logo_process = PhotoImage(master=fr_process, file=r'img/logo_process.gif')
        logo_process = Label(fr_process, borderwidth=2, image=self.logo_process)

        # options
        caz_go_excel = Checkbutton(fr_process,
                                   text=u'Exporter tout le catalogue en Excel',
                                   variable=self.opt_excel)

        caz_go_word = Checkbutton(fr_process,
                                   text=u'Exporter chaque métadonnée en Word',
                                   variable=self.opt_word)
        
        # launcher
        self.btn_go = Button(fr_process,
                             text="Lancer l'export",
                             command=lambda: process_wordification())

        # griding widgets
        logo_process.grid(row=1, rowspan=3,
                          column=0, padx=2,
                          pady=2, sticky="W")
        
        caz_go_word.grid(row=2, column=1)
        caz_go_excel.grid(row=2, column=2)
        self.btn_go.grid(row=3, column=1, columnspan=2, sticky="WE")
        
# ----------------------------------------------------------------------------

    def get_input_xl(self):
        """ Get the path of the input Excel file with a browse dialog
        """
        self.input_xl = askopenfilename(parent=self,
                                        filetypes=[("Excel 2010 files","*.xlsx"),("Excel 2003 files","*.xls")],
                                        title=u"Choisir le fichier Excel à partir duquel faire la jointure")

        # testing file choosen
        if self.input_xl:
            print(self.input_xl)
            pass
        elif path.splittext(self.input_xl)[1] != ".xlsx":
            print("Pas le bon format")
        else:
            print(u'Aucun fichier sélectionné')
            return

        # get headers names
        xlsx_in = load_workbook(filename=self.input_xl,
                                read_only=True,
                                guess_types=True,
                                use_iterators=True)
        ws1 = xlsx_in.worksheets[0]  # ws = première feuille
        cols_names = [ws1.cell(row=ws1.min_row, column=col).value for col in range(1, ws1.max_column)]

        # end of method
        return

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

    def open_urls(self, li_url):
        """ Open URLs in new tabs in the default brower.
        It waits a few seconds between the first and the next URLs
        to handle case when the webbrowser is not opened yet and let the
        time to do.
        """
        x = 1
        for url in li_url:        
            if x > 1: 
                sleep()
            else:
                pass
            open_new_tab(url)
            x += 1

        # end of method
        return


# ----------------------------------------------------------------------------

    def settings_load(self, config_file=r"settings.ini"):
        """ TO DO
        """
        config = SafeConfigParser()
        config.read(r"settings.ini")
        self.settings = {s:dict(config.items(s)) for s in config.sections()}

        logger.info("Settings loaded from: {}".format(config_file))

        # end of method
        return

    def settings_save(self):
        """ TO DO
        """

        logger.info("Settings saved into: {}".format(config_file))
        # end of method
        return

# ----------------------------------------------------------------------------

    def get_shares_info(self):
        """TO DOCUMENT
        """
        # variables
        li_oc = []
        li_owners = []
        li_without_oc = []
        # parsing
        for share in self.shares:
            # Share caracteristics
            share_name = share.get("name")
            creator_name = share.get("_creator").get("contact").get("name")
            creator_id = share.get("_creator").get("_tag")[6:]
            share_url = "https://app.isogeo.com/groups/{}/admin/shares/{}"\
                        .format(creator_id, share.get("_id"))

            li_owners.append(creator_id)    # add to shares owners list
            # OpenCatalog URL construction
            share_details = self.isogeo.share(self.token, share_id=share.get("_id"))
            url_OC = "http://open.isogeo.com/s/{}/{}".format(share.get("_id"),
                                                             share_details.get("urlToken"))

            # Testing URL
            request = requests.get(url_OC)
            if request.status_code != 200:
                logger.info("No OpenCatalog set for this share: " + share_url)
                li_without_oc.append(share_url)
                continue
            else:
                pass
            
            # consolidate list of OpenCatalog available
            li_oc.append((share_name, creator_id, creator_name, share_url, url_OC))
            
        # end of method
        return li_oc, set(li_owners), li_without_oc

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

# ----------------------------------------------------------------------------

    def process_excelization(self, output_filename):
        """ TO DO
        """
        includes = ["conditions",
                    "contacts",
                    "coordinate-system",
                    "events",
                    "feature-attributes",
                    "keywords",
                    "limitations",
                    "links",
                    "specifications"]

        self.search_results = self.isogeo.search(self.token,
                                                 page_size=0,
                                                 whole_share=0)

        # ------------ REAL START ----------------------------
        wb = Isogeo2xlsx()
        wb.set_worksheets()

        # parsing metadata
        for md in search_results.get('results'):
            wb.store_metadatas(md)

        # tunning
        wb.tunning_worksheets()

        # saving the test file
        dstamp = datetime.now()
        wb.save(r"output\{0}.xlsx".format())

        # end of method
        return

    def process_wordification(self, search_results):
        """ TO DO
        """
        for md in search_results.get("results"):
            tpl = DocxTemplate(path.realpath(self.tpl_input.get()))
            toDocx.md2docx(tpl, md, url_oc)
            dstamp = datetime.now()
            if not md.get('name'):
                md_name = "NR"
            elif '.' in md.get('name'):
                md_name = md.get("name").split(".")[1]
            else:
                md_name = md.get("name")
            tpl.save(r"..\output\{0}_{8}_{7}_{1}{2}{3}{4}{5}{6}.docx".format("TestDemoDev",
                                                                             dstamp.year,
                                                                             dstamp.month,
                                                                             dstamp.day,
                                                                             dstamp.hour,
                                                                             dstamp.minute,
                                                                             dstamp.second,
                                                                             md.get("_id")[:5],
                                                                             md_name))
            del tpl

        # end of method
        return

# ----------------------------------------------------------------------------

    def no_ui_launcher(self):
        """ Execute the scripts without displaying the UI and using
        settings.ini
        """
        logger.info('Launched from command prompt')
        self.settings_load()
        exit()
        pass

# ###############################################################################
# ###### Stand alone program ########
# ###################################

if __name__ == '__main__':
    """ standalone execution
    """
    if len(argv) < 2:
        app = Isogeo2office(ui_launcher=1)
        app.mainloop()
    elif argv[1] == str(1):
        print("launch UI")
        app = Isogeo2office(ui_launcher=1)
        app.mainloop()
    else:
        print("launch without UI")
        app = Isogeo2office(ui_launcher=0)
