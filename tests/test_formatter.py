# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)

# #############################################################################
# ########## Libraries #############
# ##################################

# Standard library
from os import environ
import logging
from random import randint
from sys import exit
import unittest

# 3rd party
from isogeo_pysdk import Isogeo, __version__ as pysdk_version

# target
from modules import IsogeoFormatter

# #############################################################################
# ######## Globals #################
# ##################################

# API access
app_id = environ.get('ISOGEO_API_DEV_ID')
app_secret = environ.get('ISOGEO_API_DEV_SECRET')

# #############################################################################
# ########## Classes ###############
# ##################################


class TestSearch(unittest.TestCase):
    """Test search to Isogeo API."""
    if not app_id or not app_secret:
        logging.critical("No API credentials set as env variables.")
        exit()
    else:
        pass
    logging.debug('Isogeo PySDK version: {0}'.format(pysdk_version))

    # standard methods
    def setUp(self):
        """Executed before each test."""
        self.isogeo = Isogeo(client_id=app_id,
                             client_secret=app_secret,
                             platform="qa")
        self.bearer = self.isogeo.connect()
        self.fmt = IsogeoFormatter()

    def tearDown(self):
        """Executed after each test."""
        pass

    # formatter
    def test_cgus(self):
        """CGU formatter."""
        search = self.isogeo.search(self.bearer, page_size=0, whole_share=0)
        licenses = [t for t in search.get("tags") if t.startswith("license:")]
        # filtered search
        md_cgu = self.isogeo.search(self.bearer,
                                    query=licenses[0],
                                    include=["conditions", ],
                                    page_size=1,
                                    whole_share=0)
        # get conditions reformatted
        cgus_in = md_cgu.get("results")[0].get("conditions", [])
        cgus_out = self.fmt.conditions(cgus_in)
        # test
        self.assertIsInstance(cgus_out, list)

    def test_limitations(self):
        """Limitations formatter."""
        search = self.isogeo.search(self.bearer,
                                    whole_share=1,
                                    include=["limitations", ])
        # filtered search
        for md in search.get("results"):
            if md.get("limitations"):
                md_lims = md
                break

        # get limitations reformatted
        lims_in = md_lims.get("limitations", [])
        lims_out = self.fmt.limitations(lims_in)
        # test
        self.assertIsInstance(lims_out, list)


# ##############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == '__main__':
    unittest.main()
