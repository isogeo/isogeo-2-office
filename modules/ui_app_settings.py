# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)
# ----------------------------------------------------------------------------
# Name:         Isogeo API minimal auth ui
# Purpose:      Minimal UI form to check application id/secret and save it
#
# Author:       Julien Moura (@geojulien)
#
# Python:       2.7.x
# Created:      28/07/2016
# Updated:      29/07/2016
# ---------------------------------------------------------------------------

# ############################################################################
# ########## Libraries #############
# ##################################

# Standard library
import gettext
import logging      # log files
from os import path
from time import sleep
from Tkinter import Tk, StringVar, HORIZONTAL
from ttk import Label, Button, Entry, Separator
from webbrowser import open_new_tab

# 3rd party library
from isogeo_pysdk import Isogeo

# ############################################################################
# ########## Classes ###############
# ##################################


class IsogeoAppAuth(Tk):
    """UI class to configure client id/secret of an Isogeo 3rd party app."""

    def __init__(self, prev_id="app_id", prev_secret="app_secret",
                 app_name="Isogeo Application", lang=None):
        """UI class to insert client id/secret of an Isogeo 3rd party application.

        keyword arguments:
            - prev_id: an eventual previous client ID to insert
            - prev_secret: an eventual previous client secret to insert
        """
        Tk.__init__(self)  # instanciating

        # localization
        try:
            # lang.install(unicode=1)
            _ = lang.gettext
            logging.info(u"Custom language set: {}"
                         .format(_(u"English").decode("UTF8")))
        except Exception, e:
            logging.error(e)
            _ = gettext.gettext
            logging.info(u"Default language set: English")

        # basics
        self.title(_(u"{} - API authentication settings").format(app_name))
        try:
            self.iconbitmap(path.dirname(__file__) + r'/../img/settings.ico')
        except:
            logging.error(u"Icon file not reachable")
            pass
        self.resizable(width=False, height=False)
        self.focus_force()

        # variables
        self.app_id = StringVar(self)
        self.app_secret = StringVar(self)
        self.app_id.set(prev_id)
        self.app_secret.set(prev_secret)

        self.msg_bar = StringVar(self)
        self.msg_bar.set(_(u"Insert access transmitted by Isogeo."))

        self.li_dest = []
        # form fields
        lb_input_id = Label(self,
                            text=_(u"Client id:"))
        ent_input_id = Entry(self,
                             textvariable=self.app_id,
                             width=70)

        lb_input_secret = Label(self,
                                text=_(u"Client secret:"))
        ent_input_secret = Entry(self,
                                 textvariable=self.app_secret,
                                 width=70)

        # buttons
        btn_test = Button(self,
                          text=_(u"\U0001F5F8 Check"),
                          command=lambda: self.test_connection())
        mailto = _("mailto:Isogeo%20Projects%20"
                   "<projects+isogeo2office@isogeo.com>?"
                   "subject=[Isogeo2office]%20Access request")
        btn_contact = Button(self,
                             text=_(u"\U0001F582 Request access"),
                             command=lambda: open_new_tab(mailto))

        # message
        Separator(self, orient=HORIZONTAL).grid(row=3,
                                                column=1,
                                                columnspan=2,
                                                sticky="WE")
        lb_msg = Label(self,
                       textvariable=self.msg_bar,
                       anchor='w')

        # griding widgets
        lb_input_id.grid(row=1, column=1, sticky="W")
        ent_input_id.grid(row=1, column=2, sticky="W")
        lb_input_secret.grid(row=2, column=1, sticky="W")
        ent_input_secret.grid(row=2, column=2, sticky="W")
        btn_test.grid(row=1, column=3, rowspan=2, sticky="NSWE")
        btn_contact.grid(row=3, column=3, rowspan=2, sticky="NSE")
        lb_msg.grid(row=4, column=1, columnspan=2, sticky="WE")

        logging.info("API form launched")

    def test_connection(self):
        """Check parameters entered."""
        try:
            self.isogeo = Isogeo(client_id=self.app_id.get(),
                                 client_secret=self.app_secret.get())
            self.token = self.isogeo.connect()
            self.msg_bar.set(_(u"Everything is fine."))
            sleep(2)
            self.li_dest = [self.app_id.get(), self.app_secret.get()]
            logging.info("New access id/secret granted")
            self.destroy()
        except Exception, e:
            logging.error(e)
            self.msg_bar.set(e[1])

        # end of method
        return

# ###############################################################################
# ###### Stand alone program ########
# ###################################

if __name__ == '__main__':
    """ standalone execution
    """
    app = IsogeoAppAuth(prev_id="Here comes the client ID",
                        prev_secret="Here comes the client secret")
    app.mainloop()
    print("New oAuth2 parameters: ", app.li_dest)
