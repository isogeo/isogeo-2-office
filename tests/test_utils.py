# -*- coding: UTF-8 -*-
#! python3

"""
    Usage from the repo root folder:
    
    ```python
    python -m unittest tests.test_utils
    # or specific test
    python -m unittest tests.test_utils.TestIsogeo2officeUtils.test_thumbnails_loader_complete
    ```
"""

# #############################################################################
# ########## Libraries #############
# ##################################

# Standard library
from os import path
import unittest
import xml.etree.ElementTree as ET

# module target
from modules import isogeo2office_utils

# #############################################################################
# ########## Classes ###############
# ##################################


class TestIsogeo2officeUtils(unittest.TestCase):
    """Test utils functions of Isogeo to Office."""

    # standard methods
    def setUp(self):
        # """Executed before each test."""
        self.utils = isogeo2office_utils()
        self.fixtures_dir = path.normpath(r"tests/fixtures")
        # thumbnails tables
        self.thumbs_complete = path.normpath(
            path.join(self.fixtures_dir, "thumbnails_complete.xlsx")
        )
        self.thumbs_bad_sheetname = path.normpath(
            path.join(self.fixtures_dir, "thumbnails_bad_worksheetName.xlsx")
        )
        self.thumbs_bad_header = path.normpath(
            path.join(self.fixtures_dir, "thumbnails_bad_headers.xlsx")
        )

    def tearDown(self):
        """Executed after each test."""
        pass

    #  -- Openers ------------------------------------------------------------
    # def test_url_opener(self):
    #     """Test URL opener"""
    #     self.utils.open_urls(["https://example.com"])
    #     self.utils.open_urls(["https://example.com", "https://example.org"])

    def test_dirfile_opener_ok(self):
        """Test file/folder opener"""
        self.utils.open_dir_file(r"modules")

    def test_dirfile_opener_bad(self):
        """Test file/folder opener"""
        with self.assertRaises(IOError):
            self.utils.open_dir_file(r"toto")

    #  -- Cleaners ------------------------------------------------------------
    def test_clean_accents(self):
        """Test special characters remover"""
        # set
        in_str = "Spécial $#! caractères n'ont que des esp@ces à droite 888323"
        # run
        clean_stripped_accent = self.utils.clean_special_chars(in_str)
        clean_underscored_accent = self.utils.clean_special_chars(in_str, "_")
        clean_stripped_pure = self.utils.clean_special_chars(in_str, accents=0)
        clean_underscored_pure = self.utils.clean_special_chars(in_str, "_", accents=0)
        # check
        self.assertEqual(
            clean_stripped_accent, "Spécialcaractèresnontquedesespcesàdroite888323"
        )
        self.assertEqual(
            clean_underscored_accent,
            "Spécial_caractères_n_ont_que_des_esp_ces_à_droite_888323",
        )
        self.assertEqual(
            clean_stripped_pure, "Spcialcaractresnontquedesespcesdroite888323"
        )
        self.assertEqual(
            clean_underscored_pure,
            "Sp_cial_caract_res_n_ont_que_des_esp_ces_droite_888323",
        )

    def test_clean_filename_ok(self):
        """Test clean filenames"""
        # set
        in_filename = "mon rapport de catalogage super cool ! .zip"
        # run
        filename_soft = self.utils.clean_filename(in_filename, mode="soft")
        filename_strict = self.utils.clean_filename(in_filename, mode="strict")
        # check
        self.assertEqual(filename_soft, "mon rapport de catalogage super cool ! .zip")
        self.assertEqual(filename_strict, "mon rapport de catalogage super cool  .zip")

    def test_clean_filename_bad(self):
        """Test filenames errors"""
        with self.assertRaises(ValueError):
            self.utils.clean_filename(r"toto", mode="youpi")

    def test_clean_xml(self):
        """Test XML cleaner"""
        # set
        in_xml = """<field name="id">abcdef</field>
                    <field name="intro" > pqrst</field>
                    <field name="desc"> this is a test file. We will show 5>2 and 3<5 and
                    try to remove non xml compatible characters.</field>
                 """
        # run
        clean_xml_soft = self.utils.clean_xml(in_xml)
        clean_xml_strict = self.utils.clean_xml(in_xml, mode="strict")
        # check
        ET.fromstring("<root>{}</root>".format(clean_xml_soft))
        ET.fromstring("<root>{}</root>".format(clean_xml_strict))

    #  -- Thumbnails ----------------------------------------------------------
    def test_thumbnails_loader_complete(self):
        """Test filenames errors"""
        expected_dict = {"1234569732454beca1ab3ec1958ffa50": "resources/table.svg"}
        self.assertDictEqual(
            self.utils.thumbnails_mngr(self.thumbs_complete), expected_dict
        )

    def test_thumbnails_loader_bad_notTable(self):
        """Test filenames errors"""
        thumbs_loaded = self.utils.thumbnails_mngr(r"table_thumbnails.xlsx")
        self.assertEqual(thumbs_loaded, {None: (None, None)})

    def test_thumbnails_loader_bad_sheetname(self):
        """Test filenames errors"""
        with self.assertRaises(KeyError):
            self.utils.thumbnails_mngr(self.thumbs_bad_sheetname)

    def test_thumbnails_loader_bad_headers(self):
        """Test filenames errors"""
        with self.assertRaises(KeyError):
            self.utils.thumbnails_mngr(self.thumbs_bad_header)
