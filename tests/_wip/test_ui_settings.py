# -*- coding: UTF-8 -*-
#! python3

"""
    Usage from the repo root folder:
    
    ```python
    python -m unittest tests.test_ui_settings
    ```
"""

# #############################################################################
# ########## Libraries #############
# ##################################

# standard
import sys
import unittest

# PyQt5
from PyQt5.QtCore import QSettings
from PyQt5.QtWidgets import QApplication
from PyQt5 import QtTest

# module target
import IsogeoToOffice

app = QApplication(sys.argv)

# #############################################################################
# ########## Classes ###############
# ##################################
class TestUiSettings(unittest.TestCase):
    """Test IsogeoToOffice QSettings management."""

    # standard methods
    def setUp(self):
        """Executed before each test."""
        self.i2o = IsogeoToOffice.IsogeoToOffice_Main()

    def tearDown(self):
        """Executed after each test."""
        pass

    #  -- Tests ------------------------------------------------------------
    def test_defaults(self):
        """Test the GUI in its default state"""
        self.assertListEqual(self.i2o.app_settings.childGroups(), ["formats", "usage"])


# #############################################################################
# ##### Main #######################
# ##################################
if __name__ == "__main__":
    unittest.main()
