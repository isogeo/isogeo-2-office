# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (print_function, unicode_literals)
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
import json
import locale
from math import ceil
from os import listdir, path
from sys import exit
from Tkinter import Tk, StringVar, IntVar    # GUI
from ttk import Label, Button, Entry, Combobox    # widgets

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
        super(Isogeo, self).__init__()
        self.arg = _lang

    def connect(self, identifant, token):
        pass

    def share(self, ):
        pass

    def catalog(self, ):
        pass

    def resource(self, id_resource):
        pass


# ###############################################################################
# ###### Stand alone program ########
# ###################################

if __name__ == '__main__':
    """ standalone execution """
    app = Isogeo2Office()
    app.mainloop()
