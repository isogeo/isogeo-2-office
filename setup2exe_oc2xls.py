# -*- coding: UTF-8 -*-
#!/usr/bin/env python
#-------------------------------------------------------------------------------
# Name:         oc2xls
# Purpose:      Script to transform a script into an Windows executable
#                   software. It uses py2exe.
#
# Author:       Julien Moura (https://github.com/Guts/)
#
# Python:       2.7.x
# Created:      19/12/2011
# Updated:      12/05/2014
#
# Licence:      GPL 3
#-------------------------------------------------------------------------------

################################################################################
########### Libraries #############
###################################
# Standard library
from distutils.core import setup
import ConfigParser
import numpy
import os
import py2exe
import sys

################################################################################
########## Main program ###########
###################################

# adding py2exe to the env path
sys.argv.append('py2exe')


# Specific dll for pywin module
mfcdir = r'C:\Python27\Lib\site-packages\pythonwin'
mfcfiles = [os.path.join(mfcdir, i) for i in ["mfc90.dll", "mfc90u.dll", "mfcm90.dll", "mfcm90u.dll", "Microsoft.VC90.MFC.manifest"]]

# build settings
build_options = dict(
                    build_base='setup/temp_build',
                    )

# conversion settings 
py2exe_options = dict(
                        excludes=['_ssl',  # Exclude _ssl
                                  'pyreadline', 'doctest', 'email',
                                  'optparse', 'pickle'],  # Exclude standard library
                        dll_excludes = ['MSVCP90.dll'],
                        compressed=1,  # Compress library.zip
                        optimize = 2,
                      )


setup(name="oc2xls",
      version="0.1",
      description="Get Isogeo metadat into an Excel workbook",
      author="Julien Moura",
      options={'py2exe': py2exe_options, 'build': build_options},
      windows = [
            {
            "script": "Isogeo_OpenCatalog2xls.py"                     # main script
            }
                ]
    )