# -*- coding: UTF-8 -*-
#!/usr/bin/env python
# -----------------------------------------------------------------------------
# Name:         isogeo2office_exe
# Purpose:      Script to transform isogeo2office scripts into a
#                   Windows executable. It uses py2exe.
#
# Author:       Julien Moura (https://github.com/Guts/)
#
# Python:       2.7.x
# Created:      19/12/2015
# Updated:      12/08/2016
#
# Licence:      GPL 3
# -----------------------------------------------------------------------------

# #############################################################################
# ######### Libraries #############
# #################################
# Standard library
from distutils.core import setup
from ConfigParser import SafeConfigParser
from os import path
import py2exe
import sys
from docxtpl import *

# custom modules
from isogeo2office import _version
from modules import *

# #############################################################################
# ######## Main program ###########
# #################################

# ------------ Initial settings ---------------------------------------------
config = SafeConfigParser()
config.read(path.realpath("settings_TPL.ini"))
with open(path.realpath("build\settings.ini"), 'wb') as configfile:
    config.write(configfile)

# ------------ Windows settings ---------------------------------------------
sys.argv.append('py2exe')   # add py2exe to the path

# Specific dll for pywin module
mfcdir = r'C:\Python27\Lib\site-packages\pythonwin'
mfcfiles = [path.join(mfcdir, i) for i in ["mfc90.dll",
                                           "mfc90u.dll",
                                           "mfcm90.dll",
                                           "mfcm90u.dll",
                                           "Microsoft.VC90.MFC.manifest"]]

# ------------ Build options ------------------------------------------------
build_options = dict(build_base='build/temp_build',
                     )

# conversion settings
py2exe_options = dict(excludes=['_ssl',  # Exclude _ssl
                                'pyreadline', 'doctest', 'email',
                                'optparse', 'pickle'],  # Exclude standard lib
                      includes=['lxml.etree', 'lxml._elementpath', 'gzip'],
                      dll_excludes=['MSVCP90.dll'],
                      compressed=1,  # Compress library.zip
                      optimize=2,
                      # bundle_files = 1,
                      dist_dir='build/isogeo2office_{}'.format(_version)
                      )


# ------------ APP settings ------------------------------------------------
setup(name="isogeo2office - {}".format(_version),
      version=_version,
      description="Export Isogeo metadata to desktop formats (Word, Excel...)",
      author="Julien Moura",
      url="https://bitbucket.org/isogeo/isogeo-2-office",
      license="license GPL v3.0",
      data_files=[("", ["build\settings.ini"]),
                  # images
                  ("", ["img/favicon.ico"]),
                  ("img", ["img/logo_isogeo.gif",
                           "img/favicon_isogeo.gif",
                           "img/logo_Word2013.gif",
                           "img/logo_Excel2013.gif",
                           "img/logo_process.gif"]),
                  # templates
                  ("templates", ["templates/template_Isogeo.docx"]),
                  # output
                  ("output", ["output/README.md"]),
                  ],
      options={'py2exe': py2exe_options,
               'build': build_options},
      windows=[{"script": "isogeo2office.py",  # main script
               "icon_resources": [(1, "img/favicon.ico")]  # Icone
                }
               ]
      )
