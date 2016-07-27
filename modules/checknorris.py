# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)
# -----------------------------------------------------------------------------
# Name:         Check Norris
# Purpose:      A class dedicated to perform system test to ensure another
#               program works fine
#
# Author:       Julien Moura (@geojulien)
#
# Python:       2.7.x
# Created:      22/12/2015
# Updated:      10/03/2016
# -----------------------------------------------------------------------------

# #############################################################################
# ###### Standard Libraries ########
# ##################################

import logging
from os import environ as env, path, listdir
import socket
import sys
from urllib2 import getproxies
from urllib2 import build_opener, install_opener, ProxyHandler, urlopen

# #############################################################################
# ########## Classes ###############
# ##################################


class CheckNorris(object):
    """ Check Norris never fails, always tests.
    """
    # -- ATTRIBUTES -----------------------------------------------------------

    # -- BEFORE ALL -----------------------------------------------------------

    def __init__(self):
        """ Check Norris welcomes you
        """
        super(CheckNorris, self).__init__()

    # -- 1 method, 1 check ----------------------------------------------------

    def check_gdal(self):
        """ Checks if OSGeo libs work and if GDAL_DATA is well refrenced.
        Returns:
        -- 1: GDAL_DATA already exists as environment variable
        -- 2: GDAL_DATA didn't exist as environment variable then has been added
        -- 3: GDAL_DATA didn't exist as environment variable then has been added
        """
        # GDAL install
        try:
            try:
                from osgeo import gdal
            except ImportError:
                import gdal
            logging.info('GDAL version: {}'.format(gdal.__version__))
        except:
            logging.error("GDAL is not installed or not reachable. DicoGIS is going to close.")
            return 1

        # GDAL_DATA variable
        if "GDAL_DATA" not in env.keys():
            try:
                gdal.SetConfigOption(str('GDAL_DATA'),
                                     str(path.abspath(r'data/gdal')))
                logging.info("GDAL_DATA path not found in environment variable.\
                                  DicoGIS'll use its own: "
                             + path.abspath(r'data/gdal'))
                return 2
            except:
                logging.error("Oups! Something's wrong with GDAL_DATA path.")
                return 3
        else:
            logging.info("GDAL_DATA path found in environment variable: {}.\
                         DicoGIS'll use it.".format(env.get("GDAL_DATA")))
            return 4
        # end of method
        return

    def check_arcpy(self):
        """ Checks if arcpy and which version is installed
        """
        # 3rd party libraries
        try:
            import arcpy
            esri_info = arcpy.GetInstallInfo()
            logging.info("ArcPy imported from ArcGIS {} v{} in ({})".format(
                         esri_info.get("ProductName"),
                         esri_info.get("Version"),
                         esri_info.get("InstallDir")))
            # end of method
            return True, esri_info
        except RuntimeError:
            logging.error("ArcPy is installed, but not licensed.")
            return False, "ArcGIS is installed, but not licensed."
        except ImportError:
            logging.info("ArcGIS isn't in the SYSPATH. Trying to find it automatically.")
            # checks if ArcGIS is installed
            if not path.isdir(path.join(env.get("PROGRAMFILES(x86)"), "ArcGIS"))\
               and not path.isdir(path.join(env.get("PROGRAMFILES"), "ArcGIS")):
                logging.info("ArcGIS isn't installed on this computer.")
                return False, "ArcGIS isn't installed on this computer."
            else:
                arcgis_path = path.join(env.get("PROGRAMFILES(x86)", "PROGRAMFILES"), "ArcGIS")
                pass
            logging.info("ArcGIS is installed but not well configured.")
            # path to the last version of 10 branch
            v = max([i[-1] for i in listdir(path.realpath(arcgis_path)) if "Desktop10" in i])
            arcgis_path = path.join(arcgis_path, "Desktop10.{}".format(v))
            # adding paths to the environment
            sys.path.append(path.join(arcgis_path, "arcpy"))
            sys.path.append(path.join(arcgis_path, "bin"))
            sys.path.append(path.join(arcgis_path, "ArcToolbox\Scripts"))
            try:
                import arcpy
                import site
                esri_info = arcpy.GetInstallInfo()
                logging.info("ArcGIS configuration has been fixed.")
                if hasattr(sys, 'real_prefix'):
                    # inside a venv
                    logging.info("Executing inside a virtualenv. Nice!")
                    pypacks = [p for p in sys.path if p.endswith('site-packages')][-1]
                else:
                    # using system install
                    logging.info("Executing from the main Python install.")
                    pypacks = site.getsitepackages()[1]

                # creatring pth file for future runs
                with open(path.join(pypacks, 'arcpy.pth'), 'w') as pth_arcpy:
                    pth_arcpy.write(path.realpath(path.join(arcgis_path, "arcpy")) + "\n")
                    pth_arcpy.write(path.realpath(path.join(arcgis_path, "bin")) + "\n")
                    pth_arcpy.write(path.realpath(path.join(arcgis_path, "ArcToolbox\Scripts")) + "\n")
                # end of method
                return True, esri_info
            except:
                logging.info("ArcGIS automatic configuration failed.")
                return False, "ArcGIS automatic configuration failed."
        else:
            logging.info("ArcGIS isn't installed on this computer.")
            return False, "ArcGIS isn't installed on this computer."

    def check_internet_connection(self, remote_server="www.google.com"):
        """ Checks if an internet connection is operational
        source: http://stackoverflow.com/a/20913928/2556577
        """
        try:
            # see if we can resolve the host name -- tells us if there is
            # a DNS listening
            host = socket.gethostbyname(remote_server)
            # connect to the host -- tells us if the host is actually
            # reachable
            socket.create_connection((host, 80), 2)
            logging.info("Internet connection OK.")
            return True
        except:
            logging.info("Internet connection failed.")
            pass
        # end of method
        return False

    def check_proxy(self, specific={}):
        """ Checks if proxy settings are set on the OS
        Returns:
        -- 1 when direct connection works fine
        -- 2 when direct connection fails and any proxy is set in the OS
        -- 3 and settings when direct connection fails but a proxy is set
        see: https://docs.python.org/2/library/urllib.html#urllib.getproxies
        """
        os_proxies = getproxies()
        if len(os_proxies) == 0 and self.check_internet_connection:
            logging.info("No proxy needed nor set. Direct connection works.")
            return 1
        elif len(os_proxies) == 0 and not self.check_internet_connection:
            logging.error("Proxy not set in the OS. Needs to be specified")
            return 2
        else:
            #
            env['http_proxy'] = os_proxies.get("http")
            env['https_proxy'] = os_proxies.get("https")
            #
            proxy = ProxyHandler({
                                 'http': os_proxies.get("http"),
                                 'https': os_proxies.get("https")
                                 })
            opener = build_opener(proxy)
            install_opener(opener)
            urlopen('http://www.google.com')
            return 3, os_proxies


# ##############################################################################
# ##### Stand alone program ########
# ##################################

if __name__ == '__main__':
    """ standalone execution """
    # ------------ Specific imports ----------------

    # ------------ Real start ----------------
    # instanciating the class
    checker = CheckNorris()

    # checking GDAL installation and configuration
    print("GDAL: ", checker.check_gdal())

    # checking arcpy installation
    print("ArcPy: ", checker.check_arcpy())

    # checking internet connection
    print("Internet: ", checker.check_internet_connection())

    # checking proxy settings
    print("Proxy: ", checker.check_proxy())
