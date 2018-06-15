# -*- coding: utf-8 -*-

# #############################################################################
# ########## Libraries #############
# ##################################

from tkinter import END
from tkinter.ttk import Entry
from tests import BaseWidgetTest

# module target
from modules import isogeo2office_utils

# #############################################################################
# ########## Classes ###############
# ##################################

class TestEntryValidators(BaseWidgetTest):
    def test_entry_validator_date(self):
        # utils methods
        self.utils = isogeo2office_utils()
        # validator
        fields_validators = {
            "val_date": (self.window.register(self.utils.entry_validate_date),
                         '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W')
        }
        # date validator
        ent_date = Entry(self.window, width=2, validate="key",
                         validatecommand=fields_validators.get("val_date")
                         )
        ent_date.pack()

        # test accepted values
        values_ok = "012"
        for i in values_ok:
            ent_date.insert(0, i)
            self.assertEqual(ent_date.get(), str(i))
            ent_date.delete(0, END)
        # test rejected values
        ent_date.insert(0, "hello")
        self.assertEqual(ent_date.get(), "")

    def test_entry_validator_uid(self):
        # utils methods
        self.utils = isogeo2office_utils()
        # validator
        fields_validators = {
            "val_uid": (self.window.register(self.utils.entry_validate_uid),
                        '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W'),
                        }
        # uid validator
        ent_uid = Entry(self.window, width=2, validate="key",
                        validatecommand=fields_validators.get("val_uid")
                        )
        ent_uid.pack()
        self.window.update()

        # test accepted values
        values_ok = "012345678"
        for i in values_ok:
            ent_uid.insert(0, i)
            self.assertEqual(ent_uid.get(), str(i))
            ent_uid.delete(0, END)
        # test rejected values
        values_bad = ["hello", "9", "10"]
        for i in values_bad:
            ent_uid.insert(0, i)
            self.assertEqual(ent_uid.get(), "")
            ent_uid.delete(0, END)
