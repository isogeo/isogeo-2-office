# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)
# ----------------------------------------------------------------------------
# Name:         isogeo2office useful methods
# Purpose:      externalize util methods from isogeo2office
#
# Author:       Julien Moura (@geojulien)
#
# Python:       3.5.x
# Created:      14/10/2017
# Updated:      14/10/2017
# ----------------------------------------------------------------------------

# ############################################################################
# ########## Libraries #############
# ##################################

# Standard library
from datetime import date, datetime
import logging
from os import path
import sqlite3

# ##############################################################################
# ############ Globals ############
# #################################

logger = logging.getLogger("isogeo2office")  # LOG

# ############################################################################
# ########## Classes ###############
# ##################################


class DbManager(object):
    """isogeo2office DB manager."""

    def __init__(self, main_path=r"."):
        """Instanciating method."""
        super(DbManager, self).__init__()

        # ------------ VARIABLES ---------------------
        self.mpath = path.realpath(main_path)
        self.db_path = path.join(self.mpath, 'db\isogeo_metadata.db')

    def get_db_connection(self):
        """Check if a DB has already been created and create one if not."""
        if path.isfile(self.db_path):
            logger.debug("DB already exists")
            return sqlite3.connect(path.join(self.mpath, 'isogeo_metadata.db'))
        else:
            return self.create_db()

    def create_db(self):
        """Init database with CREATE instructions."""
        logger.info("Creating the new DB")
        # importing SQL script
        with open('init.sql', 'r') as in_sql_file:
            sql = in_sql_file.read()
        # execute SQL script
        with sqlite3.connect(self.db_path) as conn:
            c = conn.cursor()
            c.executescript(sql)
        conn.commit()
        return sqlite3.connect(self.db_path)

    def update_metadata(self, api_results):
        """Store metadata results into the database."""
        with sqlite3.connect(self.db_path) as conn:
            c = conn.cursor()
            for i in api_results:
                now = datetime.now()
                c.execute("INSERT INTO metadata VALUES ({}, {}, {}, {}, {}, {}, {}, {}, {}, {})".format(i.get("_id"),
                                                                                                        i.get("title"),
                                                                                                        i.get("abstract"),
                                                                                                        i.get("_created"),
                                                                                                        now,
                                                                                                        now,
                                                                                                        i.get("_modified"),
                                                                                                        i.get("created", now),
                                                                                                        i.get("modified", now),
                                                                                                        1,))
                conn.commit()

        pass

# ############################################################################
# ##### Stand alone program ########
# ##################################

if __name__ == '__main__':
    """Standalone execution and tests"""
    # specific imports
    import json
    # starting
    db_mngr = DbManager(r"..\..")
    db_mngr.get_db_connection()
    # import sample API results
    with open(r'..\..\tests\out_api_search_basic.json') as data_file:
        data = json.load(data_file)
    # print(type(data), data.keys())
    db_mngr.update_metadata(data.get("results"))
