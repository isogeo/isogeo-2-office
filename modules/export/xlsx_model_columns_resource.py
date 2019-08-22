# -*- coding: UTF-8 -*-
#! python3

"""
    Matching table between Isogeo metadata model and Excle columns for Isogeo to Office.
"""

RESOURCE_COLUMNS = {
    "_created": "T",
    "_creator": "D",
    "_id": "S",
    "_modified": "U",
    "abstract": "B",
    "conditions": "K",
    "contacts": "M",
    "created": "F",
    "format": "J",
    "keywords": "E",
    "language": "V",
    "limitations": "L",
    "links": "",
    "modified": "H",
    "name": "Z",
    "path": "C",
    "published": "I",
    "specifications": "AB",
    "tags": "",
    "title": "A",
    # specific
    "hasLinkDownload": "N",
    "hasLinkOther": "P",
    "hasLinkView": "O",
    "linkEdit": "Q",
    "linkView": "R",
    "inspireConformance": "Y",
    "inspireThemes": "X",
}
