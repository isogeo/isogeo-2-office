# -*- coding: UTF-8 -*-
#!/usr/bin/env python

"""
    Usage from the repo root folder:
    
    ```python
    python -m unittest tests.test_export_docx
    ```
"""

# #############################################################################
# ########## Libraries #############
# ##################################

# Standard library
import json
from os import environ, path
from tempfile import mkstemp
import unittest

# 3rd party
from isogeo_pysdk import Isogeo, IsogeoTranslator
from docxtpl import DocxTemplate

# target
from modules import Isogeo2docx
from modules import IsogeoFormatter
from modules import IsogeoStats

# #############################################################################
# ######## Globals #################
# ##################################

# API access
app_id = environ.get('ISOGEO_API_DEV_ID')
app_secret = environ.get('ISOGEO_API_DEV_SECRET')

# #############################################################################
# ########## Classes ###############
# ##################################


class TestExportDocx(unittest.TestCase):
    """Test export to Microsoft Word DOCX."""
    def setUp(self):
        """Executed before each test."""
        # API response samples
        self.search_all_includes = path.normpath(
            r"tests/fixtures/api_search_complete.json")
        # template
        self.word_template = path.normpath(
            r"templates/template_Isogeo.docx")

        # target class instanciation
        self.to_docx = Isogeo2docx()

    def tearDown(self):
        """Executed after each test."""
        pass

    # -- Export ---------------------------------------------------------------
    def test_metadata_export(self):
        """Test search results export"""
        # temp output file
        #out_docx = mkstemp(prefix="i2o_test_docx_")
        # load tags fixtures
        with open(self.search_all_includes, "r") as f:
            search = json.loads(f.read())
        # load template
        tpl = DocxTemplate(self.word_template)
        url_oc = "https://open.isogeo.com/"
        # run
        for md in search.get('results'):
            # output path
            out_docx = mkstemp(prefix="i2o_test_docx_")
            out_docx_path = out_docx[1] + ".docx"
            # templating
            tpl = DocxTemplate(self.word_template)
            self.to_docx.md2docx(tpl, md, url_oc)
            # save
            tpl.save(out_docx_path)
            del tpl


# #############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == '__main__':
    unittest.main()
