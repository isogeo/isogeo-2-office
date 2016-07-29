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
import logging      # log files
from time import sleep
from Tkinter import Tk, StringVar
from ttk import Label, Button, Entry
from webbrowser import open_new_tab

# 3rd party library
from isogeo_pysdk import Isogeo

# custom
from checknorris import CheckNorris

# ############################################################################
# ########## Classes ###############
# ##################################


class IsogeoAppAuth(Tk):
    """UI Class to
    docstring for Isogeo to Office
    """

    def __init__(self, prev_id="app_id", prev_secret="app_secret"):
        """ TO DOC
        """
        # Invoke Check Norris
        checker = CheckNorris()

        # checking connection
        if not checker.check_internet_connection():
            logging.error('Internet connection required. Check your settings.')
            exit()
        else:
            pass

        # instanciating
        Tk.__init__(self)

        # basics
        self.title(u'isogeo2office - Paramètres')
        self.iconbitmap('../img/settings.ico')
        self.resizable(width=False, height=False)
        self.focus_force()

        # variables
        self.app_id = StringVar(self)
        self.app_secret = StringVar(self)
        self.app_id.set(prev_id)
        self.app_secret.set(prev_secret)

        self.msg_bar = StringVar(self)
        self.msg_bar.set("Insérer les informations d'accès transmises par Isogeo.")

        self.li_dest = []
        # form fields
        lb_input_id = Label(self,
                            text="Client id :")
        ent_input_id = Entry(self,
                             textvariable=self.app_id,
                             width=70)

        lb_input_secret = Label(self,
                                text="Client secret :")
        ent_input_secret = Entry(self,
                                 textvariable=self.app_secret,
                                 width=70)

        # test button
        btn_test = Button(self,
                          text="Check",
                          command=lambda: self.test_connection())

        # message
        lb_msg = Label(self,
                       textvariable=self.msg_bar)

        # griding widgets
        lb_input_id.grid(row=1, column=1, sticky="W")
        ent_input_id.grid(row=1, column=2, sticky="W")
        lb_input_secret.grid(row=2, column=1, sticky="W")
        ent_input_secret.grid(row=2, column=2, sticky="W")
        btn_test.grid(row=1, column=3, rowspan=2, sticky="NSE")
        lb_msg.grid(row=3, columnspan=3, sticky="WE")

    def test_connection(self):
        """TODOC
        """
        try:
            self.isogeo = Isogeo(client_id=self.app_id.get(),
                                 client_secret=self.app_secret.get())
            self.token = self.isogeo.connect()
            self.msg_bar.set("Tout est ok.")
            sleep(2)
            self.li_dest = [self.app_id.get(), self.app_secret.get()]
            self.destroy()
        except Exception, e:
            self.msg_bar.set(e[1])

        # end of method
        return


# ###############################################################################
# ###### Stand alone program ########
# ###################################

if __name__ == '__main__':
    """ standalone execution
    """
    app = IsogeoAppAuth()
    app.mainloop()
    print("New access: ", app.li_dest)
