# -*- coding: UTF-8 -*-
#! python3

"""
    Matching table between Isogeo metadata model and Excle columns for Isogeo to Office.
"""

VECTOR_COLUMNS = {
    # model
    "_created": "AM",
    "_creator": "E",
    "_id": "AL",
    "_modified": "AN",
    "abstract": "C",
    "collectionContext": "I",
    "collectionMethod": "J",
    "conditions": "AD",
    "contacts": "AF",
    "coordinateSystem": "T",
    "created": "O",
    "distance": "W",
    "editionProfile": "",
    "encoding": "",
    "envelope": "U",
    "events": "P",  # count
    "featureAttributes": "AA",
    "features": "Y",
    "format": "S",
    "formatVersion": "",
    "geometry": "V",
    "keywords": "F",
    "language": "AO",
    "layers": "",
    "limitations": "AE",
    "links": "",
    "modified": "Q",
    "name": "B",
    "operations": "",
    "path": "D",
    "precision": "",
    "published": "R",
    "scale": "X",
    "series": "",
    "serviceLayers": "",
    "specifications": "AB",
    "tags": "",
    "title": "A",
    "topologicalConsistency": "AC",
    "type": "",
    "updateFrequency": "M",
    "validFrom": "K",
    "validTo": "L",
    "validityComment": "N",
    # specific
    "featureAttributesCount": "Z",
    "hasLinkDownload": "AG",
    "hasLinkOther": "AI",
    "hasLinkView": "AH",
    "linkEdit": "AJ",
    "linkView": "AK",
    "inspireConformance": "H",
    "inspireThemes": "G",
}
