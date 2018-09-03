# -*- coding: UTF-8 -*-
#!/usr/bin/env python

# #############################################################################
# ########## Libraries #############
# ##################################

# Standard library
from collections import namedtuple
import logging
from os import environ, path
from six import string_types as str
from sys import exit
from tempfile import mkstemp
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
        self.ini_file = path.normpath(r"tests/fixtures/settings_TPL.ini")
        self.ini_out = mkstemp(prefix="i2o_test_settings_")

    def tearDown(self):
        """Executed after each test."""
        pass

    #  -- Openers ------------------------------------------------------------
    def test_url_opener(self):
        """Test URL opener"""
        self.utils.open_urls(["https://example.com", ])
        self.utils.open_urls(["https://example.com",
                              "https://example.org"])

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
        self.assertEqual(clean_stripped_accent,
                         "Spécialcaractèresnontquedesespcesàdroite888323")
        self.assertEqual(clean_underscored_accent,
                         "Spécial_caractères_n_ont_que_des_esp_ces_à_droite_888323")
        self.assertEqual(clean_stripped_pure,
                         "Spcialcaractresnontquedesespcesdroite888323")
        self.assertEqual(clean_underscored_pure,
                         "Sp_cial_caract_res_n_ont_que_des_esp_ces_droite_888323")

    def test_clean_filename_ok(self):
        """Test clean filenames"""
        # set
        in_filename = "mon rapport de catalogage super cool ! .zip"
        # run
        filename_soft = self.utils.clean_filename(in_filename, mode="soft")
        filename_strict = self.utils.clean_filename(in_filename, mode="strict")
        # check
        self.assertEqual(filename_soft,
                         "mon rapport de catalogage super cool ! .zip")
        self.assertEqual(filename_strict,
                         "mon rapport de catalogage super cool  .zip")

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


    #  -- Settings manager ----------------------------------------------------
    def test_settings_loader(self):
        """test settings loader"""
        # run
        settings = self.utils.settings_load(config_file=self.ini_file)

        # check settings structure
        self.assertIsInstance(settings, dict)
        self.assertIn("auth", settings)
        self.assertIn("local", settings)
        self.assertIn("proxy", settings)
        self.assertIn("excel", settings)
        self.assertIn("word", settings)
        self.assertIn("xml", settings)

        # check auth section
        self.assertIn("app_id", settings.get("auth"))
        self.assertIn("app_secret", settings.get("auth"))
        self.assertEqual(settings.get("auth").get("app_id"), 
                         "python-minimalist-sdk-test-uuid-1a2b3c4d5e6f7g8h9i0j11k12l")
        self.assertEqual(settings.get("auth").get("app_secret"),
                         "application-secret-1a2b3c4d5e6f7g8h9i0j11k12l13m14n15o16p17Q18rS")

        # check proxy section
        self.assertIn("proxy_needed", settings.get("proxy"))
        self.assertIn("proxy_type", settings.get("proxy"))
        self.assertIn("proxy_prot", settings.get("proxy"))
        self.assertIn("proxy_server", settings.get("proxy"))
        self.assertIn("proxy_port", settings.get("proxy"))
        self.assertIn("proxy_user", settings.get("proxy"))

        # check excel section
        self.assertIn("opt_attributes", settings.get("excel"))
        self.assertIn("opt_fillfull", settings.get("excel"))
        self.assertIn("excel_opt", settings.get("excel"))
        self.assertIn("opt_inspire", settings.get("excel"))
        self.assertIn("output_name", settings.get("excel"))

        # check word section
        self.assertIn("word_opt", settings.get("word"))
        self.assertIn("out_prefix", settings.get("word"))
        self.assertIn("tpl", settings.get("word"))
        self.assertIn("opt_date", settings.get("word"))
        self.assertIn("opt_id", settings.get("word"))

        # check xml section
        self.assertIn("out_prefix", settings.get("xml"))
        self.assertIn("xml_opt", settings.get("xml"))
        self.assertIn("opt_zip", settings.get("xml"))
        self.assertIn("opt_date", settings.get("xml"))
        self.assertIn("opt_id", settings.get("xml"))
