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

    # def md2wb(self, wbsheet, offset, li_mds, li_catalogs):
    #     """
    #     parses Isogeo metadatas and write it into the worksheet
    #     """
    #     # looping on metadata
    #     for md in li_mds:
    #         # incrementing line number
    #         offset += 1
    #         # extracting & parsing tags
    #         tags = md.get("tags")
    #         li_motscles = []
    #         li_theminspire = []
    #         srs = ""
    #         owner = ""
    #         inspire_valid = 0
    #         # looping on tags
    #         for tag in tags.keys():
    #             # free keywords
    #             if tag.startswith('keyword:isogeo'):
    #                 li_motscles.append(tags.get(tag))
    #                 continue
    #             else:
    #                 pass
    #             # INSPIRE themes
    #             if tag.startswith('keyword:inspire-theme'):
    #                 li_theminspire.append(tags.get(tag))
    #                 continue
    #             else:
    #                 pass
    #             # workgroup which owns the metadata
    #             if tag.startswith('owner'):
    #                 owner = tags.get(tag)
    #                 continue
    #             else:
    #                 pass
    #             # coordinate system
    #             if tag.startswith('coordinate-system'):
    #                 srs = tags.get(tag)
    #                 continue
    #             else:
    #                 pass
    #             # format pretty print
    #             if tag.startswith('format:'):
    #                 format_lbl = tags.get(tag)
    #                 continue
    #             else:
    #                 format_lbl = "NR"
    #                 pass
    #             # INSPIRE conformity
    #             if tag.startswith('conformity:inspire'):
    #                 inspire_valid = 1
    #                 continue
    #             else:
    #                 pass

    #         # HISTORY ###########
    #         if md.get("created"):
    #             data_created = dtparse(md.get("created")).strftime("%a %d %B %Y")
    #         else:
    #             data_created = "NR"
    #         if md.get("modified"):
    #             data_updated = dtparse(md.get("modified")).strftime("%a %d %B %Y")
    #         else:
    #             data_updated = "NR"
    #         if md.get("published"):
    #             data_published = dtparse(md.get("published")).strftime("%a %d %B %Y")
    #         else:
    #             data_published = "NR"

    #         # formatting links to visualize on OpenCatalog and edit on APP
    #         link_visu = 'HYPERLINK("{0}"; "{1}")'.format(url_OpenCatalog + "/m/" + md.get('_id'),
    #                                                      "Visualiser")
    #         link_edit = 'HYPERLINK("{0}"; "{1}")'.format("https://app.isogeo.com/resources/" + md.get('_id'),
    #                                                      "Editer")
    #         # format version
    #         if md.get("formatVersion"):
    #             format_version = u"{0} ({1} - {2})".format(format_lbl,
    #                                                        md.get("formatVersion"),
    #                                                        md.get("encoding"))
    #         else:
    #             format_version = format_lbl

    #         # formatting contact details
    #         contacts = md.get("contacts")
    #         if len(contacts):
    #             contacts_cct = ["{0} ({1}) ;\n".format(contact.get("contact").get("name"),
    #                                                    contact.get("contact").get("email"))\
    #                             for contact in contacts if contact.get("role") == "pointOfContact"]
    #         else:
    #             contacts_cct = ""

    #         # METADATA #
    #         md_created = dtparse(md.get("_created")).strftime("%a %d %B %Y (%Hh%M)")
    #         md_updated = dtparse(md.get("_modified")).strftime("%a %d %B %Y (%Hh%M)")

    #         # écriture des informations dans chaque colonne correspondante
    #         wbsheet.write(offset, 0, md.get("title"))
    #         wbsheet.write(offset, 1, md.get("name"))
    #         wbsheet.write(offset, 2, md.get("path"))
    #         wbsheet.write(offset, 3, " ; ".join(li_motscles))
    #         wbsheet.write(offset, 4, md.get("abstract"), style_wrap)
    #         wbsheet.write(offset, 5, " ; ".join(li_theminspire))
    #         wbsheet.write(offset, 6, md.get("type"))
    #         wbsheet.write(offset, 7, format_version)
    #         wbsheet.write(offset, 8, srs)
    #         wbsheet.write(offset, 9, md.get("features"))
    #         wbsheet.write(offset, 10, md.get("geometry"))
    #         wbsheet.write(offset, 11, owner)
    #         wbsheet.write(offset, 12, data_created.decode('latin1'))
    #         wbsheet.write(offset, 13, data_updated.decode('latin1'))
    #         wbsheet.write(offset, 14, md_created.decode('latin1'))
    #         wbsheet.write(offset, 15, md_updated.decode('latin1'))
    #         wbsheet.write(offset, 16, inspire_valid)
    #         wbsheet.write(offset, 17, len(contacts))
    #         wbsheet.write(offset, 18, contacts_cct, style_wrap)
    #         wbsheet.write(offset, 20, xlwt.Formula(link_visu), style_url)
    #         wbsheet.write(offset, 21, xlwt.Formula(link_edit), style_url)

    #     print(sorted(md.keys()))

    #     # end of function
    #     return

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
        else:
            self.ws_v["D{}".format(self.idx_v)] = src_path
            pass

        # owner
        self.ws_v["E{}".format(self.idx_v)] = next(v for k, v in tags.items() if 'owner:' in k)

        # KEYWORDS
        if md.get("keywords"):
            self.ws_v["F{}".format(self.idx_v)] = " ; ".join([k.get("text") for k in md.get("keywords")])
        else:
            self.ws_v["F{}".format(self.idx_v)] = ""

        # INSPIRE
        if md.get("inspire-theme"):
            self.ws_v["G{}".format(self.idx_v)] = " ; ".join([k.get("text") for k in md.get("inspire-theme")])
        else:
            self.ws_v["G{}".format(self.idx_v)] = ""

        # conformity
        self.ws_v["H{}".format(self.idx_v)] = "conformity:inspire" in tags

        # HISTORY
        self.ws_v["I{}".format(self.idx_v)] = md.get("collectionContext", "")
        self.ws_v["J{}".format(self.idx_v)] = md.get("collectionMethod", "")
        self.ws_v["K{}".format(self.idx_v)] = md.get("validFrom", "")
        self.ws_v["L{}".format(self.idx_v)] = md.get("validTo", "")
        self.ws_v["M{}".format(self.idx_v)] = md.get("updateFrequency", "")
        self.ws_v["N{}".format(self.idx_v)] = md.get("validComment", "")

        # EVENTS
        # data creation date
        if md.get("created"):
            data_created = arrow.get(md.get("created"))
            data_created = "{0} ({1})".format(data_created.format("DD/MM/YYYY", "fr_FR"),
                                              data_created.humanize(locale="fr_FR"))
        else:
            data_created = "NR"
        self.ws_v["O{}".format(self.idx_v)] = data_created

        # events count
        self.ws_v["P{}".format(self.idx_v)] = len(md.get('events', ""))

        # data last update
        if md.get("modified"):
            data_updated = arrow.get(md.get("created"))
            data_updated = "{0} ({1})".format(data_updated.format("DD/MM/YYYY", "fr_FR"),
                                              data_updated.humanize(locale="fr_FR"))
        else:
            data_updated = "NR"
        self.ws_v["Q{}".format(self.idx_v)] = data_updated

        # TECHNICAL
        # format
        if 'format' in md.keys():
            format_lbl = next(v for k, v in tags.items() if 'format:' in k)
        else:
            format_lbl = "NR"
        self.ws_v["S{}".format(self.idx_v)] = u"{0} ({1} - {2})".format(format_lbl,
                                                                        md.get("formatVersion", "NR"),
                                                                        md.get("encoding", "NR"))

        # SRS
        srs = md.setdefault("coordinate-system", {"name":"NR", "code": "NR"})
        self.ws_v["T{}".format(self.idx_v)] = u"{0} ({1})".format(srs.get("name", "NR"),
                                                                  srs.get("code", "NR"))

        # LINKS
        link_edit = r'=HYPERLINK("{0}","{1}")'.format("https://app.isogeo.com/resources/" + md.get("_id"),
                                                      "Visualiser")
        self.ws_v["AL{}".format(self.idx_v)] = link_edit
        self.ws_v["AL{}".format(self.idx_v)].style = self.s_link

        # STYLING
        self.ws_v["C{}".format(self.idx_v)].style = self.s_wrap

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
    from datetime import datetime

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
    includes = ["coordinate-system", "events", "links"]

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
