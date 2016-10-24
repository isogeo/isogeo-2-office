# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)
# ----------------------------------------------------------------------------
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
import gettext  # localization
import logging
from logging.handlers import RotatingFileHandler
from os import listdir, path
import platform  # about operating systems
from sys import argv, exit
from time import sleep
from tkFileDialog import askdirectory, askopenfilename
from tkMessageBox import showerror as avert
from Tkinter import Tk, Image, PhotoImage
from Tkinter import IntVar, StringVar, ACTIVE, DISABLED, VERTICAL
from ttk import Label, Button, Entry, Checkbutton, Combobox
from ttk import Labelframe, Progressbar, Separator, Style
from webbrowser import open_new_tab

# 3rd party library
from docxtpl import DocxTemplate
from isogeo_pysdk import Isogeo
from openpyxl import load_workbook
import requests

# Custom modules
from modules.isogeo2xlsx import Isogeo2xlsx
from modules.isogeo2docx import Isogeo2docx
from modules.ui_app_settings import IsogeoAppAuth
from modules.checknorris import CheckNorris

# ############################################################################
# ########## Global ################
# ##################################

# VERSION
_version = "1.5.3"

# LOG FILE ##
# see: http://sametmax.com/ecrire-des-logs-en-python/
logger = logging.getLogger()
logging.captureWarnings(True)
logger.setLevel(logging.INFO)  # all errors will be get
log_form = logging.Formatter("%(asctime)s || %(levelname)s "
                             "|| %(module)s || %(message)s")
logfile = RotatingFileHandler("LOG_isogeo2office.log", "a", 5000000, 1)
logfile.setLevel(logging.INFO)
logfile.setFormatter(log_form)
logger.addHandler(logfile)
logger.info('=================================================')
logger.info('================ Isogeo => Office ===============')


# ############################################################################
# ########## Classes ###############
# ##################################

class Isogeo2office(Tk):
    """Main Class for Isogeo to Office."""

    # attributes and global actions
    logging.info('OS: {0}'.format(platform.platform()))
    logging.info('Version: {0}'.format(_version))

    def __init__(self, ui_launcher=1):
        """Initiliazing isogeo2office with or without UI."""
        # Invoke Check Norris
        checker = CheckNorris()

        # checking connection
        if not checker.check_internet_connection():
            logger.error('Internet connection required: check your settings.')
            exit()
        else:
            pass

        # UI or not to UI
        if not ui_launcher:
            self.no_ui_launcher()
        else:
            pass

        # ------------ Settings ----------------------------------------------
        self.settings_load()
        self.app_id = self.settings.get("auth").get("app_id")
        self.app_secret = self.settings.get('auth').get("app_secret")
        self.client_lang = self.settings.get('basics').get("def_codelang")

        # ------------ Localization ------------------------------------------
        if self.client_lang == "FR":
            lang = gettext.translation("isogeo2office", localedir="i18n",
                                       languages=["fr_FR"], codeset="Latin1")
            lang.install(unicode=1)
        else:
            lang = gettext
            lang.install("isogeo2office", localedir="i18n",
                         unicode=1)
            pass
        logger.info("Language applied: {}".format(_("English")))

        # ------------ Isogeo authentication -------------------------------
        try:
            self.isogeo = Isogeo(client_id=self.app_id,
                                 client_secret=self.app_secret,
                                 lang=self.client_lang)
            self.token = self.isogeo.connect()
        except:
            # if id/secret doesn't work, ask for a new one
            prompter = IsogeoAppAuth(prev_id=self.app_id,
                                     prev_secret=self.app_secret,
                                     lang=lang)
            prompter.mainloop()
            # check response
            if len(prompter.li_dest) < 2:
                logger.error(u"API authentication form returned nothing.")
                exit()
            else:
                pass

            self.app_id = prompter.li_dest[0]
            self.app_secret = prompter.li_dest[1]
            self.isogeo = Isogeo(client_id=self.app_id,
                                 client_secret=self.app_secret,
                                 lang=self.client_lang)
            self.token = self.isogeo.connect()

        # ------------ Isogeo search & shares --------------------------------
        self.search_results = self.isogeo.search(self.token,
                                                 page_size=0,
                                                 whole_share=0)
        self.shares = self.isogeo.shares(self.token)
        self.shares_info = self.get_shares_info()

        # ------------ Variables ---------------------------------------------
        li_tpls = [path.abspath(path.join(r'templates', tpl))
                   for tpl in listdir(r'templates')
                   if path.splitext(tpl)[1].lower() == ".docx"]

        # ------------ UI ----------------------------------------------------
        Tk.__init__(self)
        self.title("isogeo2office - {0}".format(_version))
        icon = Image("photo", file=r"img/favicon_isogeo.gif")
        self.call("wm", "iconphoto", self._w, icon)
        self.style = Style(self).theme_use("vista")
        self.resizable(width=False, height=False)
        self.focus_force()
        self.msg_bar = StringVar(self)
        self.msg_bar.set(_(u"Pick your options and push the launch button"))

        # styling
        btn_style_err = Style(self)
        btn_style_err.configure('Error.TButton', foreground='Red')

        cbb_style_err = Style(self)
        cbb_style_err.configure('TCombobox', foreground='Red')

        # Frames and main widgets
        fr_isogeo = Labelframe(self, name="isogeo", text="Isogeo")
        fr_excel = Labelframe(self, name="excel", text="Excel")
        fr_word = Labelframe(self, name="word", text="Word")
        fr_process = Labelframe(self, name="process", text="Launch")
        self.status_bar = Label(self, textvariable=self.msg_bar, anchor='w',
                                foreground='DodgerBlue')
        self.progbar = Progressbar(self,
                                   orient="horizontal")

        fr_isogeo.grid(row=1, column=1, padx=2, pady=4, sticky="WE")
        fr_excel.grid(row=2, column=1, padx=2, pady=4, sticky="WE")
        fr_word.grid(row=3, column=1, padx=2, pady=4, sticky="WE")
        fr_process.grid(row=4, column=1, padx=2, pady=4, sticky="WE")
        self.status_bar.grid(row=5, column=1, padx=2, pady=2, sticky="WE")
        self.progbar.grid(row=6, column=1, sticky="WE")

        # --------------------------------------------------------------------

        # ## GLOBAL ##
        self.app_metrics = StringVar(fr_isogeo)
        self.oc_msg = StringVar(fr_isogeo)
        self.url_input = StringVar(fr_isogeo)

        # logo
        self.logo_isogeo = PhotoImage(master=fr_isogeo,
                                      file=r'img/logo_isogeo.gif')
        logo_isogeo = Label(fr_isogeo, borderwidth=2, image=self.logo_isogeo)

        # metrics
        self.app_metrics.set(_("{} metadata in\n"
                               "{} shares owned by\n"
                               "{} workgroups.")
                             .format(self.search_results.get('total'),
                                     len(self.shares),
                                     len(self.shares_info[1])))
        lb_app_metrics = Label(fr_isogeo,
                               textvariable=self.app_metrics)

        # OpenCatalog check
        self.lb_input_oc = Label(fr_isogeo,
                                 textvariable=self.oc_msg)

        if len(self.shares_info[2]) != 0:
            logger.error("Any OpenCatalog found among the shares")
            self.oc_msg.set(_("{} shares don't have any OpenCatalog."
                              "\nAdd OpenCatalog to every share,"
                              "\nthen reboot isogeo2office.").format(len(self.shares_info[2])))
            self.msg_bar.set(_("Error: some shares don't have OpenCatalog"
                               "activated. Fix it first."))
            self.status_bar.config(foreground='Red')
            btn_open_shares = Button(fr_isogeo,
                                     text=_("Fix the shares"),
                                     command=lambda: self.open_urls(self.shares_info[2]))
            status_launch = DISABLED
        elif len(self.shares) != len(self.shares_info[1]):
            logger.error("More than one share by workgroup")
            self.oc_msg.set(_("Too much shares by workgroup."
                              "\nPlease red button to fix it"
                              "\nthen reboot isogeo2office.")
                            .format(len(self.shares) - len(self.shares_info[1])))
            self.msg_bar.set(_("Error: more than one share by worgroup."
                               " Click on Admin button to fix it."))
            self.status_bar.config(foreground='Red')
            btn_open_shares = Button(fr_isogeo,
                                     text=_("Fix the shares"),
                                     command=lambda: self.open_urls(self.shares_info[3]),
                                     style="Error.TButton")
            status_launch = DISABLED
        else:
            logger.info("All shares have an OpenCatalog")
            self.oc_msg.set(_("Configuration OK."))
            li_oc = [share[3] for share in self.shares_info[0]]
            btn_open_shares = Button(fr_isogeo,
                                     text="\U0001F6E0 " + _("Admin shares"),
                                     command=lambda: self.open_urls(li_oc))
            status_launch = ACTIVE

        # settings
        btn_settings = Button(fr_isogeo,
                              text="\U0001F510 " + _("Settings"),
                              command=lambda: self.ui_settings_prompt())

        # contact
        mailto = _("mailto:Isogeo%20Projects%20"
                   "<projects+isogeo2office@isogeo.com>?"
                   "subject=[Isogeo2office]%20Question")
        btn_contact = Button(fr_isogeo,
                             text="\U0001F582 " + _("Contact"),
                             command=lambda: open_new_tab(mailto))

        # source
        url_src = "https://bitbucket.org/isogeo/isogeo-2-office"
        btn_src = Button(fr_isogeo,
                         text="\U0001F56C " + _("Report"),
                         command=lambda: open_new_tab(url_src))

        # griding widgets
        logo_isogeo.grid(row=1, rowspan=3,
                         column=0, padx=2,
                         pady=2, sticky="W")
        Separator(fr_isogeo, orient=VERTICAL).grid(row=1, rowspan=3,
                                                   column=1, padx=2,
                                                   pady=2, sticky="NSE")
        lb_app_metrics.grid(row=1, column=2, rowspan=3, sticky="NWE")
        self.lb_input_oc.grid(row=2, column=2, sticky="WE")
        btn_open_shares.grid(row=1, rowspan=1,
                             column=3, padx=2, pady=2,
                             sticky="NWE")
        btn_settings.grid(row=2, rowspan=1,
                          column=3, padx=2, pady=2,
                          sticky="NWE")
        btn_contact.grid(row=1, rowspan=1,
                         column=4, padx=2, pady=2,
                         sticky="NWE")
        btn_src.grid(row=2, rowspan=1,
                     column=4, padx=2, pady=2,
                     sticky="NWE")

        # --------------------------------------------------------------------

        # ## EXCEL ##
        # variables
        self.output_xl = StringVar(fr_excel, self.settings.get("basics")
                                                          .get("excel_out",
                                                               "isogeo2xlsx"))
        # self.opt_xl_join = IntVar(fr_excel)
        # self.input_xl_join_col = StringVar(fr_excel)
        # self.input_xl = ""
        # li_input_xl_cols = []

        # logo
        self.logo_excel = PhotoImage(master=fr_excel,
                                     file=r'img/logo_excel2013.gif')
        logo_excel = Label(fr_excel, borderwidth=2, image=self.logo_excel)\

        # output file
        lb_output_xl = Label(fr_excel,
                             text=_("Output filename: "))
        ent_output_xl = Entry(fr_excel,
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
        Separator(fr_excel, orient=VERTICAL).grid(row=1, rowspan=3,
                                                  column=1, padx=2,
                                                  pady=2, sticky="NSE")
        lb_output_xl.grid(row=2, column=2, sticky="W")
        ent_output_xl.grid(row=2, column=3, sticky="WE")

        # --------------------------------------------------------------------

        # ## WORD ##
        # variables
        self.tpl_input = StringVar(fr_word)
        self.out_word_prefix = StringVar(fr_word, self.settings.get("basics")
                                                      .get("word_out_prefix",
                                                           "isogeo2docx"))
        self.word_opt_id = IntVar(fr_word, self.settings.get("basics")
                                               .get("word_opt_id", 5))
        self.word_opt_date = IntVar(fr_word, self.settings.get("basics")
                                                 .get("word_opt_date", 1))

        val_uid = (self.register(self.entry_validate_uid),
                   '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W')
        val_date = (self.register(self.entry_validate_date),
                    '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W')

        # logo
        self.logo_word = PhotoImage(master=fr_word,
                                    file=r'img/logo_word2013.gif')
        logo_word = Label(fr_word, borderwidth=2, image=self.logo_word)

        # pick a template
        lb_input_tpl = Label(fr_word,
                             text=_("Pick a template: "))
        cb_available_tpl = Combobox(fr_word,
                                    textvariable=self.tpl_input,
                                    values=li_tpls)

        # specific options
        lb_out_word_prefix = Label(fr_word, text=_("File prefix: "))
        lb_out_word_uid = Label(fr_word, text=_("UID chars:\n"
                                                "(0 - 8)"))
        lb_out_word_date = Label(fr_word, text=_("Timestamp:\n"
                                                 "(0=no, 1=date, 2=datetime)"))

        ent_out_word_prefix = Entry(fr_word, textvariable=self.out_word_prefix)
        ent_out_word_uid = Entry(fr_word, textvariable=self.word_opt_id,
                                 width=2, validate="key",
                                 validatecommand=val_uid)
        ent_out_word_date = Entry(fr_word, textvariable=self.word_opt_date,
                                  width=2, validate="key",
                                  validatecommand=val_date)

        # griding widgets
        logo_word.grid(row=1, rowspan=3,
                       column=0, padx=2,
                       pady=2, sticky="W")
        Separator(fr_word, orient=VERTICAL).grid(row=1, rowspan=3,
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

        # --------------------------------------------------------------------

        # ## PROCESS ##
        # variables
        self.opt_excel = IntVar(fr_process,
                                int(self.settings.get('basics')
                                    .get('excel_opt', 0))
                                )
        self.opt_word = IntVar(fr_process,
                               int(self.settings.get('basics')
                                   .get('word_opt', 0)))
        self.out_fold_path = StringVar(fr_process,
                                       path.realpath(self.settings.get('basics')
                                                    .get('out_folder',
                                                         'output')))

        # logo
        self.logo_process = PhotoImage(master=fr_process,
                                       file=r'img/logo_process.gif')
        logo_process = Label(fr_process, borderwidth=2,
                             image=self.logo_process)

        # options
        caz_go_excel = Checkbutton(fr_process,
                                   text=_(u'Export the whole shares into an Excel worksheet'),
                                   variable=self.opt_excel)

        caz_go_word = Checkbutton(fr_process,
                                  text=_(u'Export each metadata into a Word file'),
                                  variable=self.opt_word)

        # output folder
        lb_out_fold_title = Label(fr_process, text=_("Output folder: "))
        # print(path.split(self.out_fold_path.get())[1])
        lb_out_fold_var = Label(fr_process,
                                textvariable=self.out_fold_path)

        btn_out_fold_path_browse = Button(fr_process,
                                          text=u"\U0001F3AF" + _("Browse"),
                                          command=lambda: self.set_out_folder_path(self.out_fold_path.get()))

        # launcher
        self.btn_go = Button(fr_process,
                             text="\U0001F680 " + _("Launch"),
                             command=lambda: self.process(),
                             state=status_launch)

        # griding widgets
        logo_process.grid(row=1, rowspan=5,
                          column=0, padx=2,
                          pady=2, sticky="W")
        Separator(fr_process, orient=VERTICAL).grid(row=1, rowspan=5,
                                                    column=1, padx=2,
                                                    pady=2, sticky="NSE")
        caz_go_word.grid(row=2, column=2, columnspan=3, padx=2, pady=2, sticky="W")
        caz_go_excel.grid(row=3, column=2, columnspan=3, padx=2, pady=2, sticky="W")
        lb_out_fold_title.grid(row=4, column=2,
                               padx=2, pady=2, sticky="W")
        lb_out_fold_var.grid(row=4, column=3,
                             padx=2, pady=2, sticky="WE")
        btn_out_fold_path_browse.grid(row=4, column=4,
                                      padx=2, pady=2, sticky="WE")
        self.btn_go.grid(row=5, column=2, columnspan=3,
                         padx=2, pady=2, sticky="WE")

        logger.info("Main UI instanciated & displayed")

# ----------------------------------------------------------------------------

    def set_out_folder_path(self, out_folder_path="output"):
        """Open a popup to select a folder and store it."""
        self.btn_go.config(state=DISABLED)  # disable launch in the meanwhile
        foldername = askdirectory(parent=self,
                                  initialdir=path.realpath(out_folder_path),
                                  mustexist=True,
                                  title=_("Select output folder"))
        # check if a folder has been choosen
        if foldername:
            self.out_fold_path.set(path.relpath(foldername))
        else:
            avert(title=_("No folder selected"),
                  message=_("You must select an output folder."))

        self.btn_go.config(state=ACTIVE)
        # end of function
        return foldername

    def get_input_xl(self):
        """Get the path of the input Excel file with a browse dialog."""
        self.input_xl = askopenfilename(parent=self,
                                        filetypes=[("Excel 2010", "*.xlsx")],
                                        title=_(u"Pick the Excel to merge"))

        # testing file choosen
        if self.input_xl:
            print(self.input_xl)
            pass
        elif path.splittext(self.input_xl)[1] != ".xlsx":
            print("Invalid format")
        else:
            print(_(u'Any file selected'))
            return

        # get headers names
        xlsx_in = load_workbook(filename=self.input_xl,
                                read_only=True,
                                guess_types=True,
                                use_iterators=True)
        ws1 = xlsx_in.worksheets[0]  # ws = première feuille
        cols_names = [ws1.cell(row=ws1.min_row, column=col).value
                      for col in range(1, ws1.max_column)]

        # end of method
        return

    def ui_switch_xljoiner(self):
        """Enable/disable the form for input xl to join."""
        if self.opt_xl_join.get():
            self.fr_input_xl_join.pack()
        else:
            self.fr_input_xl_join.pack_forget()
        # end of function
        return

    def open_urls(self, li_url):
        """Open URLs in new tabs in the default brower.

        It waits a few seconds between the first and the next URLs
        to handle case when the webbrowser is not opened yet and let the
        time to do.
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

# ----------------------------------------------------------------------------

    def settings_load(self, config_file=r"settings.ini"):
        """Load settings from the ini file."""
        config = SafeConfigParser()
        config.read(r"settings.ini")
        self.settings = {s: dict(config.items(s)) for s in config.sections()}

        logger.info("Settings loaded from: {}".format(config_file))

        # end of method
        return

    def settings_save(self, config_file=r"settings.ini"):
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
        # writing
        with open(path.realpath(config_file), 'wb') as configfile:
            config.write(configfile)

        logger.info("Settings saved into: {}".format(config_file))
        # end of method
        return

# ----------------------------------------------------------------------------

    def ui_settings_prompt(self):
        """Get Isogeo settings from another form."""
        prompter = IsogeoAppAuth(prev_id=self.app_id,
                                 prev_secret=self.app_secret,
                                 lang=self.client_lang)
        prompter.mainloop()
        # check response
        if len(prompter.li_dest) < 2:
            logger.error(u"API authentication form returned nothing.")
            exit()
            return 0
        else:
            pass

        self.app_id = prompter.li_dest[0]
        self.app_secret = prompter.li_dest[1]

        # end of method
        return 1

    def get_shares_info(self):
        """Get Isogeo shares informations from application."""
        # variables
        li_oc = []
        li_owners = []
        li_without_oc = []
        li_too_shares = []
        # parsing
        for share in self.shares:
            # Share caracteristics
            share_name = share.get("name")
            creator_name = share.get("_creator").get("contact").get("name")
            creator_id = share.get("_creator").get("_tag")[6:]
            share_url = "https://app.isogeo.com/groups/{}/admin/shares/{}"\
                        .format(creator_id, share.get("_id"))

            # check if there is only a share per workgroup
            if creator_id in li_owners:
                logger.error("This workgroup has more than 1 share to this application:"
                             " https://app.isogeo.com/groups/{}".format(creator_id))
                li_too_shares.append(share_url)
            else:
                # add to shares owners list
                li_owners.append(creator_id)
                pass

            # OpenCatalog URL construction
            share_details = self.isogeo.share(self.token, share_id=share.get("_id"))
            url_oc = "http://open.isogeo.com/s/{}/{}".format(share.get("_id"),
                                                             share_details.get("urlToken"))

            # Testing URL
            request = requests.get(url_oc)
            if request.status_code != 200:
                logger.info("No OpenCatalog set for this share: " + share_url)
                li_without_oc.append(share_url)
                continue
            else:
                pass

            # consolidate list of OpenCatalog available
            li_oc.append((share_name, creator_id, creator_name,
                          share_url, url_oc))

        logger.info("Isogeo - Shares informations retrieved.")
        # end of method
        return li_oc, set(li_owners), li_without_oc, li_too_shares

    def get_url_base(self, url_input):
        """Get OpenCatalog base URL to add resource ID easily."""
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

    def process(self):
        """Process export according to options set."""
        if not (self.opt_excel.get() + self.opt_word.get()):
            self.msg_bar.set(_("Error: at least one export option required"))
            logger.error("Any export option selected.")
            return
        else:
            pass

        # savings in ini file
        self.settings_save()

        # prepare Isogeo request
        self.msg_bar.set(_("Fetching Isogeo data..."))
        includes = ["conditions",
                    "contacts",
                    "coordinate-system",
                    "events",
                    "feature-attributes",
                    "keywords",
                    "layers",
                    "limitations",
                    "links",
                    "operations",
                    "serviceLayers",
                    "specifications"]

        self.search_results = self.isogeo.search(self.token,
                                                 sub_resources=includes)
        self.progbar["maximum"] = self.search_results.get("total")
        logger.info("Isogeo - metadatas retrieved.")
        # export
        if self.opt_excel.get():
            logger.info("Excel - START")
            self.progbar["value"] = 0
            self.process_excelization()
        else:
            pass

        if self.opt_word.get() and path.isfile(self.tpl_input.get()):
            self.status_bar.config(foreground='DodgerBlue')
            logger.info("WORD - START")
            self.progbar["value"] = 0
            self.process_wordification()
        elif self.opt_word.get() and self.tpl_input.get() == "":
            logger.error("Any template selected.")
            self.msg_bar.set(_("Error: Word template not selected"))
            self.status_bar.config(foreground='Red')
            return
        else:
            self.status_bar.config(foreground='DodgerBlue')
            pass

        # end of method
        self.msg_bar.set(_("All tasks are done."))
        logger.info("All tasks are done.")
        return

    def process_excelization(self):
        """Export metadatas shared into an Excel worksheet."""
        # check infos required
        if self.output_xl.get() == "":
            self.output_xl.set("isogeo2xlsx")
        else:
            pass

        # worksheet
        wb = Isogeo2xlsx()
        wb.set_worksheets()

        # parsing metadata
        for md in self.search_results.get('results'):
            wb.store_metadatas(md)
            # progression
            self.msg_bar.set(_("Processing Excel: {}").format(md.get("title")))
            self.progbar["value"] = self.progbar["value"] + 1

        # tunning
        wb.tunning_worksheets()

        # saving the test file
        # dstamp = datetime.now()
        out_xlsx_path = path.realpath(path.join(self.out_fold_path.get(),
                                                self.output_xl.get() + ".xlsx"))
        wb.save(out_xlsx_path)

        logger.info("Excel - DONE {}"
                    .format(out_xlsx_path))
        self.msg_bar.set(_("Export Excel done."))
        # end of method
        return

    def process_wordification(self):
        """Export each metadata shared to a Word document.

        Transformation is based on the template selected.
        """
        # transformer
        to_docx = Isogeo2docx()

        for md in self.search_results.get("results"):
            # get OpenCatalog related to each metadata
            if len(self.shares) == 1:
                url_oc = [share[4] for share in self.shares_info[0]][0]
            elif len(self.shares) > 1:
                # for share in self.shares_info[0]:
                #     print("\nshare owner: ", share[1])
                #     print("md owner: ", md.get("_creator").get("_id"))
                # print("\n", self.shares_info[0])
                url_oc = [share[4] for share in self.shares_info[0]
                          if share[1] == md.get("_creator").get("_id")][0]
                # print(url_oc)
            elif len(self.shares) != len(self.shares_info[1]):
                logger.error("More than one share by workgroup")
                self.msg_bar.set(_("Error: more than one share by worgroup."
                                   " Open APP or push the button to fix it."))
                self.status_bar.config(foreground='Red')
                return
            else:
                self.status_bar.config(foreground='DodgerBlue')
                pass

            # templating
            tpl = DocxTemplate(path.realpath(self.tpl_input.get()))
            to_docx.md2docx(tpl, md, url_oc)

            # name
            md_name = md.get("name", "NR")
            if '.' in md_name:
                md_name = md_name.split(".")[1]
            else:
                pass
            md_name = "_{}".format(md_name)

            # uid
            if self.word_opt_id.get():
                uid = "_{}".format(md.get("_id")[:self.word_opt_id.get()])
            else:
                uid = ""

            # date
            dstamp = datetime.now()
            if self.word_opt_date.get() == 1:
                dstamp = "_{}-{}-{}".format(dstamp.year,
                                            dstamp.month,
                                            dstamp.day)
            elif self.word_opt_date.get() == 2:
                dstamp = "_{}-{}-{}-{}{}{}".format(dstamp.year,
                                                   dstamp.month,
                                                   dstamp.day,
                                                   dstamp.hour,
                                                   dstamp.minute,
                                                   dstamp.second,)
            else:
                dstamp = ""

            # final output name
            out_docx_path = path.join(path.realpath(self.out_fold_path.get()),
                                      "{}{}{}{}.docx"
                                      .format(self.out_word_prefix.get(),
                                              uid,
                                              md_name,
                                              dstamp))
            tpl.save(out_docx_path)
            del tpl

            # progression bar
            self.msg_bar.set(_("Processing Word: {}").format(md_name))
            self.progbar["value"] = self.progbar["value"] + 1
            self.update()

        self.msg_bar.set(_("Export Word done."))
        logger.info("Word - DONE: {}".format(out_docx_path))
        # end of method
        return

# ----------------------------------------------------------------------------

    def no_ui_launcher(self):
        """Execute scripts without UI and using settings.ini."""
        logger.info('Launched from command prompt')
        self.settings_load()
        exit()
        return

# ###############################################################################
# ###### Stand alone program ########
# ###################################

if __name__ == '__main__':
    """standalone execution
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
