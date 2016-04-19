# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)
# ------------------------------------------------------------------------------
# Name:         Isogeo to Microsoft Excel 2010
# Purpose:      Get metadatas from an Isogeo share and store it into
#               a Excel worksheet. It's one of the submodules of
#               isogeo2office (https://bitbucket.org/isogeo/isogeo-2-office).
#
# Author:       Julien Moura (@geojulien)
#
# Python:       2.7.x
# Created:      14/08/2014
# Updated:      15/04/2016
# ------------------------------------------------------------------------------

# ##############################################################################
# ########## Libraries #############
# ##################################

# Standard library
from datetime import datetime
import logging
from os import path

# 3rd party library
import arrow
from openpyxl import Workbook
from openpyxl.cell import get_column_letter
from openpyxl.styles import Style, Font, Alignment
from openpyxl.worksheet.properties import WorksheetProperties

# ##############################################################################
# ########## Classes ###############
# ##################################


class Isogeo2xlsx(Workbook):
    """ Used to store Isogeo API results into an Excel worksheet (.xlsx)
    """
    cols_v = [
                      "Titre",  # A
                      "Nom",  # B
                      "Résumé",  # C
                      "Emplacement",  # D
                      "Groupe de travail",  # E
                      "Mots-clés",  # F
                      "Thématique(s) INSPIRE",  # G
                      "Conformité INSPIRE",  # H
                      "Contexte de collecte",  # I
                      "Méthode de collecte",  # J
                      "Début de validité",  # K
                      "Fin de validité",  # L
                      "Fréquence de mise à jour",  # M
                      "Commentaire",  # N
                      "Création",  # O
                      "# mises à jour",  # P
                      "Dernière mise à jour",  # Q
                      "Publication",  # R
                      "Format (version - encodage)",  # S
                      "SRS (EPSG)",  # T
                      "Emprise",  # U
                      "Géométrie",  # V
                      "Résolution",  # W
                      "Echelle",  # X
                      "# Objets",  # Y
                      "# Attributs",  # Z
                      "Attributs (A-Z)",  # AA
                      "Spécifications",  # AB
                      "Cohérence topologique",  # AC
                      "Conditions",  # AD
                      "Limitations",  # AE
                      "# Contacts",  # AF
                      "Points de contact",  # AG
                      "Autres contacts",  # AH
                      "Téléchargeable",  # AI
                      "Visualisable",  # AJ
                      "Autres",  # AK
                      "Editer",  # AL
                      "Consulter",  # AM
                      "MD - ID",  # AN
                      "MD - Création",  # AO
                      "MD - Modification",  # AP
                      "MD - Langue",  # AQ
                     ]

    cols_r = [
                        "Titre",  # A
                        "Nom",  # B
                        "Résumé",  # C
                        "Emplacement",  # D
                        "Groupe de travail",  # E
                        "Mots-clés",  # F
                        "Thématique(s) INSPIRE",  # G
                        "Conformité INSPIRE",  # H
                        "Contexte de collecte",  # I
                        "Méthode de collecte",  # J
                        "Début de validité",  # K
                        "Fin de validité",  # L
                        "Fréquence de mise à jour",  # M
                        "Commentaire",  # N
                        "Création",  # O
                        "# mises à jour",  # P
                        "Dernière mise à jour",  # Q
                        "Publication",  # R
                        "Format (version - encodage)",  # S
                        "SRS (EPSG)",  # T
                        "Emprise",  # U
                        "Résolution",  # V
                        "Echelle",  # W
                        "Attributs (A-Z)",  # X
                        "Spécifications",  # Y
                        "Cohérence topologique",  # Z
                        "Conditions",  # AA
                        "Limitations",  # AB
                        "# Contacts",  # AC
                        "Points de contact",  # AD
                        "Autres contacts",  # AE
                        "Téléchargeable",  # AF
                        "Visualisable",  # AG
                        "Autres",  # AH
                        "Editer",  # AI
                        "Consulter",  # AJ
                        "MD - ID",  # AK
                        "MD - Création",  # AL
                        "MD - Modification",  # AM
                        "MD - Langue",  # AN
                     ]

    cols_s = [
                        "Titre",  # A
                        "Nom",  # B
                        "Résumé",  # C
                        "Emplacement",  # D
                        "Groupe de travail",  # E
                        "Mots-clés",  # F
                        "Conformité INSPIRE",  # G
                        "Création",  # H
                        "# mises à jour",  # I
                        "Dernière mise à jour",  # J
                        "Publication",  # K
                        "Format (version)",  # L
                        "Emprise",  # M
                        "Spécifications",  # N
                        "Conditions",  # O
                        "Limitations",  # P
                        "# Contacts",  # Q
                        "Points de contact",  # R
                        "Autres contacts",  # S
                        "Téléchargeable",  # T
                        "Visualisable",  # U
                        "Autres",  # V
                        "Editer",  # W
                        "Consulter",  # X
                        "MD - ID",  # Y
                        "MD - Création",  # Z
                        "MD - Modification",  # AA
                        "MD - Langue",  # AB
                        ]

    cols_rz = [
                        "Titre",  # A
                        "Nom",  # B
                        "Résumé",  # C
                        "Emplacement",  # D
                        "Groupe de travail",  # E
                        "Mots-clés",  # F
                        "Création",  # G
                        "# mises à jour",  # H
                        "Dernière mise à jour",  # I
                        "Publication",  # J
                        "Format (version)",  # K
                        "Conditions",  # L
                        "Limitations",  # M
                        "# Contacts",  # N
                        "Points de contact",  # O
                        "Autres contacts",  # P
                        "Téléchargeable",  # Q
                        "Visualisable",  # R
                        "Autres",  # S
                        "Editer",  # T
                        "Consulter",  # U
                        "MD - ID",  # V
                        "MD - Création",  # W
                        "MD - Modification",  # X
                        "MD - Langue",  # Y
                        ]

    def __init__(self):
        """ Instanciating the output workbook
        """
        super(Isogeo2xlsx, self).__init__()
        # super(Isogeo2xlsx, self).__init__(write_only=True)

        # styles
        self.s_date = Style(number_format='dd/mm/yyyy')
        self.s_error = Style(font=Font(color="FF0000"))
        self.s_header = Style(alignment=Alignment(horizontal='center',
                                                  vertical='center'
                                                  ),
                              font=Font(size=12,
                                        bold=True,
                                        )
                              )
        self.s_link = Style(font=Font(underline="single"))
        self.s_wrap = Style(alignment=Alignment(wrap_text=True))
        # deleting the default worksheet
        ws = self.active
        self.remove_sheet(ws)

    # ------------ Setting workbook ---------------------

    def set_worksheets(self, vector=1, raster=1, service=1, resource=1):
        """ Adds news sheets depending on present metadata types
        """
        # SHEETS & HEADERS
        if vector:
            self.ws_v = self.create_sheet(title="Vecteurs")
            # headers
            self.ws_v.append([i for i in self.cols_v])
            # styling
            for i in self.cols_v:
                self.ws_v.cell(row=1,
                               column=self.cols_v.index(i) + 1).style = self.s_header
            # initialize line counter
            self.idx_v = 1
            # log
            logging.info("Vectors sheet added")
        else:
            pass

        if raster:
            self.ws_r = self.create_sheet(title="Raster")
            # headers
            self.ws_r.append([i for i in self.cols_r])
            # styling
            for i in self.cols_r:
                self.ws_r.cell(row=1,
                               column=self.cols_v.index(i) + 1).style = self.s_header
            # initialize line counter
            self.idx_r = 1
            # log
            logging.info("Rasters sheet added")
        else:
            pass

        if service:
            self.ws_s = self.create_sheet(title="Services")
            # headers
            self.ws_s.append([i for i in self.cols_s])
            # styling
            for i in self.cols_s:
                self.ws_s.cell(row=1,
                               column=self.cols_s.index(i) + 1).style = self.s_header
            # initialize line counter
            self.idx_s = 1
            # log
            logging.info("Services sheet added")
        else:
            pass

        if resource:
            self.ws_rz = self.create_sheet(title="Ressources")
            # headers
            self.ws_rz.append([i for i in self.cols_rz])
            # styling
            for i in self.cols_rz:
                self.ws_rz.cell(row=1,
                                column=self.cols_rz.index(i) + 1).style = self.s_header
            # initialize line counter
            self.idx_rz = 1
            # log
            logging.info("Resources sheet added")
        else:
            pass

        # end of method
        return

    # ------------ Writing metadata ---------------------

    def store_metadatas(self, metadata):
        """ TO DOCUMENT
        """
        if metadata.get("type") == "vectorDataset":
            self.idx_v += 1
            self.store_md_vector(metadata)
            return
        elif metadata.get("type") == "rasterDataset":
            self.idx_r += 1
            self.store_md_raster(metadata)
            return
        elif metadata.get("type") == "service":
            self.idx_s += 1
            self.store_md_service(metadata)
            return
        elif metadata.get("type") == "resource":
            self.idx_rz += 1
            self.store_md_resource(metadata)
            return
        else:
            print("Type of metadata is not recognized/handled: " + metadata.get("type"))
            pass
        # end of method
        return

    def store_md_vector(self, md):
        """ TO DOCUMENT
        """
        # variables
        tags = md.get("tags")

        # IDENTIFICATION
        self.ws_v["A{}".format(self.idx_v)] = md.get('title', "")
        self.ws_v["B{}".format(self.idx_v)] = md.get('name', "")
        self.ws_v["C{}".format(self.idx_v)] = md.get('abstract', "")

        # path to source
        src_path = md.get('path', "")
        if path.isfile(src_path):
            link_path = r'=HYPERLINK("{0}","{1}")'.format(path.dirname(src_path),
                                                          src_path)
            self.ws_v["D{}".format(self.idx_v)] = link_path
            logging.info("Path reachable")
        else:
            self.ws_v["D{}".format(self.idx_v)] = src_path
            logging.info("Path not recognized nor reachable")
            pass

        # owner
        self.ws_v["E{}".format(self.idx_v)] = next(v for k, v in tags.items()
                                                   if 'owner:' in k)

        # KEYWORDS & INSPIRE THEMES
        keywords = []
        inspire = []
        if "keywords" in md.keys():
            for k in md.get("keywords"):
                if k.get("_tag").startswith("keyword:is"):
                    keywords.append(k.get("text"))
                elif k.get("_tag").startswith("keyword:in"):
                    inspire.append(k.get("text"))
                else:
                    logging.info("Unknown keyword type: " + k.get("_tag"))
                    continue
            self.ws_v["F{}".format(self.idx_v)] = " ;\n".join(sorted(keywords))
            self.ws_v["G{}".format(self.idx_v)] = " ;\n".join(sorted(inspire))
        else:
            logging.info("Vector dataset without any keyword or INSPIRE theme")

        # conformity
        self.ws_v["H{}".format(self.idx_v)] = "conformity:inspire" in tags

        # HISTORY
        self.ws_v["I{}".format(self.idx_v)] = md.get("collectionContext", "")
        self.ws_v["J{}".format(self.idx_v)] = md.get("collectionMethod", "")

        # validity
        if md.get("validFrom"):
            valid_start = arrow.get(md.get("validFrom"))
            valid_start = "{0}".format(valid_start.format("DD/MM/YYYY", "fr_FR"))
        else:
            valid_start = ""
        self.ws_v["K{}".format(self.idx_v)] = valid_start

        if md.get("validTo"):
            valid_end = arrow.get(md.get("validTo"))
            valid_end = "{0}".format(valid_end.format("DD/MM/YYYY", "fr_FR"))
        else:
            valid_end = ""
        self.ws_v["L{}".format(self.idx_v)] = valid_end

        self.ws_v["M{}".format(self.idx_v)] = md.get("updateFrequency", "")
        self.ws_v["N{}".format(self.idx_v)] = md.get("validComment", "")

        # EVENTS
        # data creation date
        if md.get("created"):
            data_created = arrow.get(md.get("created"))
            data_created = "{0} ({1})".format(data_created.format("DD/MM/YYYY",
                                                                  "fr_FR"),
                                              data_created.humanize(locale="fr_FR"))
        else:
            data_created = ""
        self.ws_v["O{}".format(self.idx_v)] = data_created

        # events count
        self.ws_v["P{}".format(self.idx_v)] = len(md.get('events', ""))

        # data last update
        if md.get("modified"):
            data_updated = arrow.get(md.get("created"))
            data_updated = "{0} ({1})".format(data_updated.format("DD/MM/YYYY",
                                                                  "fr_FR"),
                                              data_updated.humanize(locale="fr_FR"))
        else:
            data_updated = ""
        self.ws_v["Q{}".format(self.idx_v)] = data_updated

        # TECHNICAL
        # format
        if 'format' in md.keys():
            format_lbl = next(v for k, v in tags.items() if 'format:' in k)
        else:
            format_lbl = ""
        self.ws_v["S{}".format(self.idx_v)] = u"{0} ({1} - {2})".format(format_lbl,
                                                                        md.get("formatVersion", "NR"),
                                                                        md.get("encoding", "NR"))

        # SRS
        srs = md.setdefault("coordinate-system", {"name": "NR", "code": "NR"})
        self.ws_v["T{}".format(self.idx_v)] = u"{0} ({1})".format(srs.get("name", "NR"),
                                                                  srs.get("code", "NR"))

        # bounding box
        bbox = md.get("envelope", None)
        if bbox:
            coords = bbox.get("coordinates")
            if bbox.get("type") == "Polygon":
                bbox = "{}\n{}".format(coords[0][0], coords[0][-2])
            elif bbox.get("type") == "Point":
                bbox = "Centroïde : {}{}".format(coords[0], coords[1])
            else:
                bbox = "Unknown envelope type (no point nor polygon): " + bbox.get("type")
        else:
            logging.info("Vector dataset without envelope.")
            pass
        self.ws_v["U{}".format(self.idx_v)] = bbox

        # geometry
        self.ws_v["V{}".format(self.idx_v)] = md.get("geometry")

        # resolution
        self.ws_v["W{}".format(self.idx_v)] = md.get("distance")

        # scale
        self.ws_v["X{}".format(self.idx_v)] = md.get("scale")

        # features objects
        self.ws_v["Y{}".format(self.idx_v)] = md.get("features")

        # features attributes
        fields = md.get("feature-attributes", None)
        if fields:
            # count
            self.ws_v["Z{}".format(self.idx_v)] = len(fields)
            # alphabetic list
            fields_cct = sorted(["{0} ({1})".format(field.get("name"),
                                                    "description" in field.keys())
                                for field in fields])
            self.ws_v["AA{}".format(self.idx_v)] = " ;\n".join(fields_cct)
        else:
            logging.info("Vector dataset without any feature attribute")
            pass

        # QUALITY
        specs = md.get("specifications", None)
        if specs:
            specs_cct = sorted(["{0} ({1})".format(s.get("specification").get("name"),
                                                   s.get("conformant"))
                                for s in specs])
            self.ws_v["AB{}".format(self.idx_v)] = " ;\n".join(specs_cct)
        else:
            logging.info("Vector dataset without specification.")
            pass
        # topology
        self.ws_v["AC{}".format(self.idx_v)] = md.get("topologicalConsistency", "")

        # CGUs
        # conditions
        conds = md.get("conditions", None)
        if conds:
            conds_cct = sorted(["{0}".format(c.setdefault("license", {"name": "No license"}).get("name"))
                               for c in conds])
            self.ws_v["AD{}".format(self.idx_v)] = " ;\n".join(conds_cct)
        else:
            logging.info("Vector dataset without conditions.")
            pass

        # limitations
        limits = md.get("limitations", None)
        if limits:
            limits_cct = sorted(["{0} ({1}) {2}".format(l.get("type"),
                                                        "{}".format(l.get("restriction", "NR")),
                                                        "{}".format(l.get("directive", {"name": ""}).get("name")))
                                for l in limits])
            self.ws_v["AE{}".format(self.idx_v)] = " ;\n".join(limits_cct)
        else:
            logging.info("Vector dataset without limitation")
            pass

        # CONTACTS
        contacts = md.get("contacts")
        if len(contacts):
            contacts_pt_cct = ["{0} ({1})".format(contact.get("contact").get("name"),
                                                  contact.get("contact").get("email"))\
                               for contact in contacts if contact.get("role") == "pointOfContact"]
            contacts_other_cct = ["{0} ({1})".format(contact.get("contact").get("name"),
                                                     contact.get("contact").get("email"))\
                                  for contact in contacts if contact.get("role") != "pointOfContact"]
            self.ws_v["AF{}".format(self.idx_v)] = len(contacts)
            self.ws_v["AG{}".format(self.idx_v)] = " ;\n".join(contacts_pt_cct)
            self.ws_v["AH{}".format(self.idx_v)] = " ;\n".join(contacts_other_cct)
        else:
            self.ws_v["AF{}".format(self.idx_v)] = 0
            logging.info("Vector dataset without any contact")
            contacts_cct = ""

        # ACTIONS
        self.ws_v["AI{}".format(self.idx_v)] = "action:download" in tags
        self.ws_v["AJ{}".format(self.idx_v)] = "action:view" in tags
        self.ws_v["AK{}".format(self.idx_v)] = "action:other" in tags

        # METADATA
        # id
        self.ws_v["AN{}".format(self.idx_v)] = md.get("_id")

        # creation
        md_created = arrow.get(md.get("_created")[:19])
        md_created = "{0} ({1})".format(md_created.format("DD/MM/YYYY",
                                                          "fr_FR"),
                                        md_created.humanize(locale="fr_FR"))
        self.ws_v["AO{}".format(self.idx_v)] = md_created

        # last update
        md_updated = arrow.get(md.get("_modified")[:19])
        md_updated = "{0} ({1})".format(md_updated.format("DD/MM/YYYY",
                                                          "fr_FR"),
                                        md_updated.humanize(locale="fr_FR"))
        self.ws_v["AP{}".format(self.idx_v)] = md_updated

        # lang
        self.ws_v["AQ{}".format(self.idx_v)] = md.get("language")

        # LINKS
        link_edit = r'=HYPERLINK("{0}","{1}")'.format("https://app.isogeo.com/resources/" + md.get("_id"),
                                                      "Editer")
        self.ws_v["AL{}".format(self.idx_v)] = link_edit
        self.ws_v["AL{}".format(self.idx_v)].style = self.s_link

        # STYLING
        self.ws_v["C{}".format(self.idx_v)].style = self.s_wrap
        self.ws_v["F{}".format(self.idx_v)].style = self.s_wrap
        self.ws_v["G{}".format(self.idx_v)].style = self.s_wrap
        self.ws_v["I{}".format(self.idx_v)].style = self.s_wrap
        self.ws_v["J{}".format(self.idx_v)].style = self.s_wrap
        self.ws_v["K{}".format(self.idx_v)].style = self.s_date
        self.ws_v["L{}".format(self.idx_v)].style = self.s_date
        self.ws_v["U{}".format(self.idx_v)].style = self.s_wrap
        self.ws_v["AA{}".format(self.idx_v)].style = self.s_wrap
        self.ws_v["AB{}".format(self.idx_v)].style = self.s_wrap
        self.ws_v["AC{}".format(self.idx_v)].style = self.s_wrap
        self.ws_v["AD{}".format(self.idx_v)].style = self.s_wrap
        self.ws_v["AE{}".format(self.idx_v)].style = self.s_wrap
        self.ws_v["AG{}".format(self.idx_v)].style = self.s_wrap
        self.ws_v["AH{}".format(self.idx_v)].style = self.s_wrap

        # LOG
        logging.info("Vector metadata stored: {} ({})".format(md.get("name"),
                                                              md.get("_id")))

        # end of method
        return

    def store_md_raster(self, md):
        """ TO DOCUMENT
        """
        # variables
        tags = md.get("tags")

        self.ws_r["A{}".format(self.idx_r)] = md.get('title')
        self.ws_r["B{}".format(self.idx_r)] = md.get('name')
        self.ws_r["C{}".format(self.idx_r)] = md.get('abstract')
        self.ws_r["D{}".format(self.idx_r)] = md.get('path')
        self.ws_r["E{}".format(self.idx_r)] = md.get('owner')

        # TECHNICAL
        # format
        if 'format' in md.keys():
            format_lbl = next(v for k, v in tags.items() if 'format:' in k)
        else:
            format_lbl = "NR"
        self.ws_r["S{}".format(self.idx_r)] = u"{0} ({1} - {2})".format(format_lbl,
                                                                        md.get("formatVersion", "NR"),
                                                                        md.get("encoding", "NR"))

        # LOG
        logging.info("Raster metadata stored: {} ({})".format(md.get("name"),
                                                              md.get("_id")))

        # end of method
        return

    def store_md_service(self, md):
        """ TO DOCUMENT
        """
        # variables
        tags = md.get("tags")

        self.ws_s["A{}".format(self.idx_s)] = md.get('title')
        self.ws_s["B{}".format(self.idx_s)] = md.get('name')
        self.ws_s["C{}".format(self.idx_s)] = md.get('abstract')
        self.ws_s["D{}".format(self.idx_s)] = md.get('path')
        self.ws_s["E{}".format(self.idx_s)] = md.get('owner')

        # TECHNICAL
        # format
        if 'format' in md.keys():
            format_lbl = next(v for k, v in tags.items() if 'format:' in k)
        else:
            format_lbl = "NR"
        self.ws_s["L{}".format(self.idx_s)] = u"{0} ({1})".format(format_lbl,
                                                                  md.get("formatVersion", "NR"))

        # LOG
        logging.info("Service metadata stored: {} ({})".format(md.get("name"),
                                                               md.get("_id")))

        # end of method
        return

    def store_md_resource(self, md):
        """ TO DOCUMENT
        """
        # variables
        tags = md.get("tags")

        self.ws_rz["A{}".format(self.idx_rz)] = md.get('title')
        self.ws_rz["B{}".format(self.idx_rz)] = md.get('name')
        self.ws_rz["C{}".format(self.idx_rz)] = md.get('abstract')
        self.ws_rz["D{}".format(self.idx_rz)] = md.get('path')
        self.ws_rz["E{}".format(self.idx_rz)] = md.get('owner')

        # TECHNICAL
        # format
        if 'format:' in tags.keys():
            format_lbl = next(v for k, v in tags.items() if 'format:' in k)
        else:
            format_lbl = "NR"
        self.ws_rz["K{}".format(self.idx_rz)] = u"{0} ({1} - {2})".format(format_lbl,
                                                                          md.get("formatVersion", "NR"),
                                                                          md.get("encoding", "NR"))

        # LOG
        logging.info("Resource metadata stored: {} ({})".format(md.get("name"),
                                                                md.get("_id")))

        # end of method
        return

    # ------------ Writing metadata ---------------------

    def tunning_worksheets(self):
        """ CLEAN UP & TUNNING
        """
        for sheet in self.worksheets:
            # Freezing panes
            c_freezed = sheet['B2']
            sheet.freeze_panes = c_freezed

            # Print properties
            sheet.print_options.horizontalCentered = True
            sheet.print_options.verticalCentered = True
            sheet.page_setup.fitToWidth = 1
            sheet.page_setup.orientation = sheet.ORIENTATION_LANDSCAPE

            # Others properties
            wsprops = sheet.sheet_properties
            wsprops.filterMode = True

            # enable filters
            sheet.auto_filter.ref = str("A1:{}{}").format(get_column_letter(sheet.max_column),
                                                          sheet.max_row)
        pass

# #############################################################################
# ##### Stand alone program ########
# ##################################

if __name__ == '__main__':
    """ Standalone execution and tests
    """
    # ------------ Specific imports ---------------------
    from ConfigParser import SafeConfigParser   # to manage options.ini

    # Custom modules
    from isogeo_sdk import Isogeo

    # ------------ Settings from ini file ----------------
    if not path.isfile(path.realpath(r"..\settings.ini")):
        logging.error("To execute this script as standalone, you need to store your Isogeo application settings in a isogeo_params.ini file. You can use the template to set your own.")
        raise ValueError("isogeo_params.ini file missing.")
    else:
        pass

    config = SafeConfigParser()
    config.read(r"..\settings.ini")

    settings = {s: dict(config.items(s)) for s in config.sections()}
    app_id = settings.get('auth').get('app_id')
    app_secret = settings.get('auth').get('app_secret')
    client_lang = settings.get('basics').get('def_codelang')

    # ------------ Connecting to Isogeo API ----------------
    # instanciating the class
    isogeo = Isogeo(client_id=app_id,
                    client_secret=app_secret,
                    lang="fr")

    token = isogeo.connect()

    # ------------ Isogeo search --------------------------
    includes = ["conditions",
                "contacts",
                "coordinate-system",
                "events",
                "feature-attributes",
                "keywords",
                "limitations",
                "links",
                "specifications"]

    search_results = isogeo.search(token,
                                   sub_resources=includes,
                                   preprocess=0)

    # ------------ REAL START ----------------------------
    wb = Isogeo2xlsx()
    wb.set_worksheets()

    # parsing metadata
    for md in search_results.get('results'):
        wb.store_metadatas(md)

    # tunning
    wb.tunning_worksheets()

    # saving the test file
    dstamp = datetime.now()
    wb.save(r"..\output\test_isogeo2xlsx_{0}{1}{2}{3}{4}{5}.xlsx".format(dstamp.year,
                                                                         dstamp.month,
                                                                         dstamp.day,
                                                                         dstamp.hour,
                                                                         dstamp.minute,
                                                                         dstamp.second))


### DEV NOTES
# http://wiki.openstreetmap.org/wiki/FR:Parcourir#URL_avec_bbox
# http://wiki.openstreetmap.org/wiki/Layer_URL_parameter
# https://www.openstreetmap.org/?bbox=22.3418234%2C57.5129102%2C22.5739625%2C57.6287332&layers=H&box=yes
