# -*- coding: UTF-8 -*-
#!/usr/bin/env python

# #############################################################################
# ########## Libraries #############
# ##################################

# Standard library
import json
from os import environ, path
import unittest

# 3rd party
from isogeo_pysdk import Isogeo, IsogeoTranslator
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle, Alignment

# target
from modules import Isogeo2xlsx
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


class TestExportXLSX(unittest.TestCase):
    """Test export to Microsoft Excel XLSX."""
    def setUp(self):
        """Executed before each test."""
        # API response samples
        self.tags_sample = path.normpath(
            r"tests/fixtures/api_response_tests_tags.json")
        self.out_wb = Isogeo2xlsx(lang="FR", url_base="https://open.isogeo.com/s")

    def tearDown(self):
        """Executed after each test."""
        pass

    def test_workbook_subclass(self):
        """Test module subclass."""
        self.assertIsInstance(self.out_wb, Workbook)
        self.assertEqual(len(self.out_wb.worksheets), 0)

    def test_output_attributes(self):
        """Test output workbook attributes."""
        # attributes
        self.assertTrue(hasattr(self.out_wb, "cols_v"))
        self.assertTrue(hasattr(self.out_wb, "cols_r"))
        self.assertTrue(hasattr(self.out_wb, "cols_s"))
        self.assertTrue(hasattr(self.out_wb, "cols_rz"))
        self.assertTrue(hasattr(self.out_wb, "cols_fa"))
        self.assertTrue(hasattr(self.out_wb, "url_base"))
        self.assertTrue(hasattr(self.out_wb, "dates_fmt"))
        self.assertTrue(hasattr(self.out_wb, "locale_fmt"))
        self.assertTrue(hasattr(self.out_wb, "stats"))
        self.assertTrue(hasattr(self.out_wb, "tr"))
        self.assertTrue(hasattr(self.out_wb, "fmt"))

        # values
        self.assertEqual(self.out_wb.url_base, "https://open.isogeo.com/s")
        self.assertIsInstance(self.out_wb.fmt, IsogeoFormatter)
        self.assertIsInstance(self.out_wb.stats, IsogeoStats)

        # languages variations
        out_wb_fr = Isogeo2xlsx(lang="FR",
                                url_base="https://open.isogeo.com/s")
        self.assertEqual(out_wb_fr.dates_fmt, "DD/MM/YYYY")
        self.assertEqual(out_wb_fr.locale_fmt, "fr_FR")

        out_wb_en = Isogeo2xlsx(lang="EN",
                                url_base="https://open.isogeo.com/s")
        self.assertEqual(out_wb_en.dates_fmt, "YYYY/MM/DD")
        self.assertEqual(out_wb_en.locale_fmt, "uk_UK")

    def test_output_structure_all_types_basic(self):
        """Test output workbook worksheets with all types tags."""
        # load tags fixtures
        with open(self.tags_sample, "r") as f:
            search = json.loads(f.read())
        # run
        self.out_wb.set_worksheets(auto=search.get("tags").keys())
        self.assertEqual(len(self.out_wb.worksheets), 4)
        self.assertIn("Raster", self.out_wb.sheetnames)
        self.assertIn("Services", self.out_wb.sheetnames)
        self.assertIn("Ressources", self.out_wb.sheetnames)
        self.assertIn("Vecteurs", self.out_wb.sheetnames)

    def test_output_structure_all_types_attributes(self):
        """Test output workbook worksheets with all types tags and attributes."""
        # load tags fixtures
        with open(self.tags_sample, "r") as f:
            search = json.loads(f.read())
        # run
        self.out_wb.set_worksheets(auto=search.get("tags").keys(), attributes=1)
        self.assertEqual(len(self.out_wb.worksheets), 5)
        self.assertIn("Raster", self.out_wb.sheetnames)
        self.assertIn("Services", self.out_wb.sheetnames)
        self.assertIn("Ressources", self.out_wb.sheetnames)
        self.assertIn("Vecteurs", self.out_wb.sheetnames)
        self.assertIn("Attributs", self.out_wb.sheetnames)


# ##############################################################################
# ##### Stand alone program ########
# ##################################
if __name__ == '__main__':
    unittest.main()
