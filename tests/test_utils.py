# -*- coding: UTF-8 -*-
#!/usr/bin/env python

# #############################################################################
# ########## Libraries #############
# ##################################

# Standard library
from six import string_types as str
import unittest

# module target
from modules import isogeo2office_utils


# #############################################################################
# ########## Classes ###############
# ##################################


class Search(unittest.TestCase):
    """Test utils functions of Isogeo to Office."""

    # standard methods
    def setUp(self):
        # """Executed before each test."""
        self.utils = isogeo2office_utils()
        self.uuid_hex = "0269803d50c446b09f5060ef7fe3e22b"
        self.uuid_urn4122 = "urn:uuid:0269803d-50c4-46b0-9f50-60ef7fe3e22b"
        self.uuid_urnIsogeo = "urn:isogeo:metadata:uuid:0269803d-50c4-46b0-9f50-60ef7fe3e22b"

    def tearDown(self):
        """Executed after each test."""
        pass

    # Isogeo components versions
    def test_get_isogeo_version_api(self):
        """"""
        version_api = self.utils.get_isogeo_version(component="api")
        version_api_naive = self.utils.get_isogeo_version()
        self.assertIsInstance(version_api, str)
        self.assertIsInstance(version_api_naive, str)
        self.assertEqual(version_api, version_api_naive)

    def test_get_isogeo_version_app(self):
        """"""
        version_app = self.utils.get_isogeo_version(component="app")
        self.assertIsInstance(version_app, str)

    def test_get_isogeo_version_db(self):
        """Check res"""
        version_db = self.utils.get_isogeo_version(component="db")
        self.assertIsInstance(version_db, str)

    def test_get_isogeo_version_bad_parameter(self):
        """Raise error if component parameter is bad."""
        with self.assertRaises(ValueError):
            self.utils.get_isogeo_version(component="youpi")

    # base URLs
    def test_set_base_url(self):
        """"""
        platform, base_url = self.utils.set_base_url()
        self.assertIsInstance(platform, str)
        self.assertIsInstance(base_url, str)

    def test_set_base_url_bad_parameter(self):
        """Raise error if platform parameter is bad."""
        with self.assertRaises(ValueError):
            self.utils.set_base_url(platform="skynet")

    # UUID converter - from HEX
    def test_hex_to_hex(self):
        """Test UUID converter from HEX to HEX"""
        uuid_out = self.utils.convert_uuid(in_uuid=self.uuid_hex,
                                           mode=0)
        self.assertIsInstance(uuid_out, str)
        self.assertEqual(uuid_out, self.uuid_hex)
        self.assertNotIn(":", uuid_out)

    def test_hex_to_urn4122(self):
        """Test UUID converter from HEX to URN (RFC4122)"""
        uuid_out = self.utils.convert_uuid(in_uuid=self.uuid_hex,
                                           mode=1)
        self.assertIsInstance(uuid_out, str)
        self.assertEqual(uuid_out, self.uuid_urn4122)
        self.assertNotIn("isogeo:metadata", uuid_out)

    def test_hex_to_urnIsogeo(self):
        """Test UUID converter from HEX to URN (Isogeo style)"""
        uuid_out = self.utils.convert_uuid(in_uuid=self.uuid_hex,
                                           mode=2)
        self.assertIsInstance(uuid_out, str)
        self.assertEqual(uuid_out, self.uuid_urnIsogeo)
        self.assertIn("isogeo:metadata", uuid_out)

    # UUID converter - from URN (RFC4122)
    def test_urn4122_to_hex(self):
        """Test UUID converter from URN (RFC4122) to HEX"""
        uuid_out = self.utils.convert_uuid(in_uuid=self.uuid_urn4122,
                                           mode=0)
        self.assertIsInstance(uuid_out, str)
        self.assertEqual(uuid_out, self.uuid_hex)
        self.assertNotIn(":", uuid_out)

    def test_urn4122_to_urn4122(self):
        """Test UUID converter from URN (RFC4122) to URN (RFC4122)"""
        uuid_out = self.utils.convert_uuid(in_uuid=self.uuid_urn4122,
                                           mode=1)
        self.assertIsInstance(uuid_out, str)
        self.assertEqual(uuid_out, self.uuid_urn4122)
        self.assertNotIn("isogeo:metadata", uuid_out)

    def test_urn4122_to_urnIsogeo(self):
        """Test UUID converter from URN (RFC4122) to URN (Isogeo style)"""
        uuid_out = self.utils.convert_uuid(in_uuid=self.uuid_urn4122,
                                           mode=2)
        self.assertIsInstance(uuid_out, str)
        self.assertEqual(uuid_out, self.uuid_urnIsogeo)
        self.assertIn("isogeo:metadata", uuid_out)

    # UUID converter - from URN (Isogeo style)
    def test_urnIsogeo_to_hex(self):
        """Test UUID converter from URN (Isogeo style) to HEX"""
        uuid_out = self.utils.convert_uuid(in_uuid=self.uuid_urnIsogeo,
                                           mode=0)
        self.assertIsInstance(uuid_out, str)
        self.assertEqual(uuid_out, self.uuid_hex)
        self.assertNotIn(":", uuid_out)

    def test_urnIsogeo_to_urn4122(self):
        """Test UUID converter from URN (Isogeo style) to URN (RFC4122)"""
        uuid_out = self.utils.convert_uuid(in_uuid=self.uuid_urnIsogeo,
                                           mode=1)
        self.assertIsInstance(uuid_out, str)
        self.assertEqual(uuid_out, self.uuid_urn4122)
        self.assertNotIn("isogeo:metadata", uuid_out)

    def test_urnIsogeo_to_urnIsogeo(self):
        """Test UUID converter from URN (Isogeo style) to URN (Isogeo style)"""
        uuid_out = self.utils.convert_uuid(in_uuid=self.uuid_urnIsogeo,
                                           mode=2)
        self.assertIsInstance(uuid_out, str)
        self.assertEqual(uuid_out, self.uuid_urnIsogeo)
        self.assertIn("isogeo:metadata", uuid_out)

    # UUID converter
    def test_uuid_converter_bad_parameter(self):
        """Raise error if one parameter is bad."""
        with self.assertRaises(ValueError):
            self.utils.convert_uuid(in_uuid="oh_my_bad_i_m_not_a_correct_uuid")
        with self.assertRaises(TypeError):
            self.utils.convert_uuid(in_uuid=2)
        with self.assertRaises(TypeError):
            self.utils.convert_uuid(in_uuid="0269803d50c446b09f5060ef7fe3e22b",
                                    mode="ups_not_an_int")
            self.utils.convert_uuid(in_uuid="0269803d50c446b09f5060ef7fe3e22b",
                                    mode=3)
            self.utils.convert_uuid(in_uuid="0269803d50c446b09f5060ef7fe3e22b",
                                    mode=True)
