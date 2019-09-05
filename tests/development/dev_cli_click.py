#!/usr/bin/python3
# -*- coding: utf-8 -*-

"""
    Isogeo to Office - Dev samples

    Command-line interface with Click https://github.com/pallets/click.
"""

# standard library
from datetime import datetime
from functools import partial
import sys

# PyQt5
from PyQt5.QtCore import QSettings
from PyQt5.QtWidgets import QApplication

# 3rd party library
import click

# #############################################################################
# ##### Main #######################
# ##################################


# get settings
app = QApplication(sys.argv)
app.settings = QSettings("Isogeo", "IsogeoToOffice")
# print(app.settings.childGroups())
# print(app.settings.allKeys())

api_auth_client_id = app.settings.value("auth/app_id", None)
api_auth_client_secret = app.settings.value("auth/app_secret", None)

# cli management
@click.command()
@click.option("--count", default=1, help="Number of greetings.")
@click.option("--name", prompt="Your name", help="The person to greet.")
def hello(count, name):
    """Simple program that greets NAME for a total of COUNT times."""
    for _ in range(count):
        click.echo("Hello, %s!" % name)


# -- MAIN --------------------------
if __name__ == "__main__":
    hello()
