#!/usr/bin/python3
# -*- coding: utf-8 -*-

"""
    Isogeo to Office - Dev samples

    Multiple radio buttons (more than 2) example (pick a timestamp)
"""

# standard library
from datetime import datetime
from functools import partial
import sys

# PyQt5
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QWidget, QPushButton, QVBoxLayout, QLabel, QApplication, QRadioButton)


class Example(QWidget):
    
    def __init__(self):
        super().__init__()
        
        self.initUI()
        
        
    def initUI(self):
        # create widgets
        self.lbl = QLabel('Select a value')
        self.date_no = QRadioButton('No date')
        self.date_day = QRadioButton('Day')
        self.date_time = QRadioButton('Datetime')

        # placement
        layout = QVBoxLayout()
        layout.addWidget(self.lbl)
        layout.addWidget(self.date_no)
        layout.addWidget(self.date_day)
        layout.addWidget(self.date_time)
        self.setLayout(layout)

        # connect radio buttons
        self.date_no.toggled.connect(partial(self.radioButtonToggled, "no"))
        self.date_day.toggled.connect(partial(self.radioButtonToggled, "day"))
        self.date_time.toggled.connect(partial(self.radioButtonToggled, "datetime"))
        
        #self.setGeometry(300, 300, 250, 150)
        self.setWindowTitle('Radio buttons - Timestamp option')
        self.show()

    def radioButtonToggled(self, timestamp_option):
        dstamp = datetime.now()
        timestamps = {
            "no": "",
            "day": "_{}-{}-{}".format(dstamp.year,
                                      dstamp.month,
                                      dstamp.day),
            "datetime": "_{}-{}-{}-{}{}{}".format(dstamp.year,
                                                  dstamp.month,
                                                  dstamp.day,
                                                  dstamp.hour,
                                                  dstamp.minute,
                                                  dstamp.second)
        }
        self.lbl.setText(timestamps.get(timestamp_option))


# -- MAIN --------------------------
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
