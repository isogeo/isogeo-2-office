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
# Updated:      12/05/2016
#
# Licence:      GPL 3
# -----------------------------------------------------------------------------

# #############################################################################
# ######### Libraries #############
# #################################
# Standard library
from distutils.core import setup
import ConfigParser
import os
import py2exe
import sys

# custom modules
from isogeo2office import _version
from modules import *

# #############################################################################
# ######## Main program ###########
# #################################

# adding py2exe to the env path
sys.argv.append('py2exe')

# Specific dll for pywin module
mfcdir = r'C:\Python27\Lib\site-packages\pythonwin'
mfcfiles = [os.path.join(mfcdir, i) for i in ["mfc90.dll", "mfc90u.dll", "mfcm90.dll", "mfcm90u.dll", " Microsoft.VC90.MFC.manifest"]]

# initial settings
confile = 'settings.ini'
config = ConfigParser.RawConfigParser()
# add sections
config.add_section('basics')
# basics
config.set('basics', 'def_codelang', 'EN')
config.set('basics', 'def_rep', './')
# Writing the configuration file
with open(confile, 'wb') as configfile:
    config.write(configfile)

# build settings
build_options = dict(build_base='setup/temp_build',
                     )

# conversion settings
py2exe_options = dict(excludes=['_ssl',  # Exclude _ssl
                                'pyreadline', 'doctest', 'email',
                                'optparse', 'pickle'],  # Exclude standard lib
                      dll_excludes=['MSVCP90.dll'],
                      compressed=1,  # Compress library.zip
                      optimize=2,
                      # bundle_files = 1,
                      dist_dir='setup/isogeo2office_{}'.format(_version)
                      )


setup(name="isogeo2office",
      version=_version,
      description="Export Isogeo metadata to desktop formats (Word, Excel...)",
      author="Julien Moura",
      url="https://bitbucket.org/isogeo/isogeo-2-office",
      license="license GPL v3.0",
      data_files=[("", ["settings.ini"]),
                  # images
                  ("", ["img/favicon.ico"]),
                  ("img", ["img/logo_isogeo.gif",
                           "img/favicon_isogeo.gif",
                           "img/logo_Word2013.gif",
                           "img/logo_Excel2003.gif"]),
                  # templates
                  ("templates", ["templates/template_Isogeo.docx"]),
                  # output
                  ("output", ["output/README.md"]),
                  ],
      options={'py2exe': py2exe_options, 'build': build_options},
      windows=[
              {"script": "isogeo2office.py",  # main script
               "icon_resources": [(1, "favicon.ico")]  # Icone
               }
              ]
     )
