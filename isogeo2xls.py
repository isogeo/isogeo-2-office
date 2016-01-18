# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (print_function, unicode_literals)
# ------------------------------------------------------------------------------
# Name:         OpenCatalog to Excel
# Purpose:      Get metadatas from an Isogeo OpenCatalog and store it into
#               an Excel workbook (.xls).
#
# Author:       Julien Moura (@geojulien) & Valentin Blanlot (@bablot)
#
# Python:       2.7.x
# Created:      14/08/2014
# Updated:      12/12/2015
# ------------------------------------------------------------------------------

###############################################################################
########### Libraries #############
###################################

# Standard library
from datetime import datetime
import json
import locale
from math import ceil
from sys import exit
from Tkinter import Tk, StringVar
from ttk import Label, Button, Entry    # widgets
from urllib2 import Request, urlopen, URLError

# 3rd party library
from dateutil.parser import parse as dtparse
import xlwt

###############################################################################
########### Classes ###############
###################################

class Isogeo2xlsx(object):
    """
    docstring for Isogeo
    """
    def __init__(self, wbsheet, offset, li_mds, li_catalogs):
        """ Isogeo connection parameters

        Keyword arguments:
        client_id -- application identifier
        client_secret -- application
        secret lang -- language asked for localized tags (INSPIRE themes).
        Could be "en" [DEFAULT] or "fr".
        proxy -- to pass through the local
        proxy. Optional. Must be a dict { 'protocol':
        'http://username:password@proxy_url:port' }.\ e.g.: {'http':
        'http://martin:p4ssW0rde@10.1.68.1:5678',\ 'https':
        'http://martin:p4ssW0rde@10.1.68.1:5678'}")
        """
        super(Isogeo, self).__init__()
        self.id = client_id
        self.ct = client_secret


###############################################################################
########## Functions ##############
###################################

def md2wb(wbsheet, offset, li_mds, li_catalogs):
    """
    parses Isogeo metadatas and write it into the worksheet
    """
    # looping on metadata
    for md in li_mds:
        # incrementing line number
        offset += 1
        # extracting & parsing tags
        tags = md.get("tags")
        li_motscles = []
        li_theminspire = []
        srs = ""
        owner = ""
        inspire_valid = 0
        # looping on tags
        for tag in tags.keys():
            # free keywords
            if tag.startswith('keyword:isogeo'):
                li_motscles.append(tags.get(tag))
                continue
            else:
                pass
            # INSPIRE themes
            if tag.startswith('keyword:inspire-theme'):
                li_theminspire.append(tags.get(tag))
                continue
            else:
                pass
            # workgroup which owns the metadata
            if tag.startswith('owner'):
                owner = tags.get(tag)
                continue
            else:
                pass
            # coordinate system
            if tag.startswith('coordinate-system'):
                srs = tags.get(tag)
                continue
            else:
                pass
            # format pretty print
            if tag.startswith('format:'):
                format_lbl = tags.get(tag)
                continue
            else:
                format_lbl = "NR"
                pass
            # INSPIRE conformity
            if tag.startswith('conformity:inspire'):
                inspire_valid = 1
                continue
            else:
                pass

        # HISTORY ###########
        if md.get("created"):
            data_created = dtparse(md.get("created")).strftime("%a %d %B %Y")
        else:
            data_created = "NR"
        if md.get("modified"):
            data_updated = dtparse(md.get("modified")).strftime("%a %d %B %Y")
        else:
            data_updated = "NR"
        if md.get("published"):
            data_published = dtparse(md.get("published")).strftime("%a %d %B %Y")
        else:
            data_published = "NR"

        # formatting links to visualize on OpenCatalog and edit on APP
        link_visu = 'HYPERLINK("{0}"; "{1}")'.format(url_OpenCatalog + "/m/" + md.get('_id'),
                                                     "Visualiser")
        link_edit = 'HYPERLINK("{0}"; "{1}")'.format("https://app.isogeo.com/resources/" + md.get('_id'),
                                                     "Editer")
        # format version
        if md.get("formatVersion"):
            format_version = u"{0} ({1} - {2})".format(format_lbl,
                                                       md.get("formatVersion"),
                                                       md.get("encoding"))
        else:
            format_version = format_lbl

        # formatting contact details
        contacts = md.get("contacts")
        if len(contacts):
            contacts_cct = ["{0} ({1}) ;\n".format(contact.get("contact").get("name"),
                                                   contact.get("contact").get("email"))\
                            for contact in contacts if contact.get("role") == "pointOfContact"]
        else:
            contacts_cct = ""

        # METADATA #
        md_created = dtparse(md.get("_created")).strftime("%a %d %B %Y (%Hh%M)")
        md_updated = dtparse(md.get("_modified")).strftime("%a %d %B %Y (%Hh%M)")

        # écriture des informations dans chaque colonne correspondante
        wbsheet.write(offset, 0, md.get("title"))
        wbsheet.write(offset, 1, md.get("name"))
        wbsheet.write(offset, 2, md.get("path"))
        wbsheet.write(offset, 3, " ; ".join(li_motscles))
        wbsheet.write(offset, 4, md.get("abstract"), style_wrap)
        wbsheet.write(offset, 5, " ; ".join(li_theminspire))
        wbsheet.write(offset, 6, md.get("type"))
        wbsheet.write(offset, 7, format_version)
        wbsheet.write(offset, 8, srs)
        wbsheet.write(offset, 9, md.get("features"))
        wbsheet.write(offset, 10, md.get("geometry"))
        wbsheet.write(offset, 11, owner)
        wbsheet.write(offset, 12, data_created.decode('latin1'))
        wbsheet.write(offset, 13, data_updated.decode('latin1'))
        wbsheet.write(offset, 14, md_created.decode('latin1'))
        wbsheet.write(offset, 15, md_updated.decode('latin1'))
        wbsheet.write(offset, 16, inspire_valid)
        wbsheet.write(offset, 17, len(contacts))
        wbsheet.write(offset, 18, contacts_cct, style_wrap)
        wbsheet.write(offset, 20, xlwt.Formula(link_visu), style_url)
        wbsheet.write(offset, 21, xlwt.Formula(link_edit), style_url)

    print(sorted(md.keys()))

    # end of function
    return

###############################################################################
######### Main program ############
###################################

# locale
locale.setlocale(locale.LC_ALL, str("fra_fra"))

##################### UI
app = Tk()
app.title('OpenCatalog ===> Excel')

# variables
url_input = StringVar(app)
lang = "fr"
start = 0

# étiquette
lb_invite = Label(app, text="Colle ici ton OpenCatalog")
lb_invite.pack()

# champ pour l'URL
ent_OpenCatalog = Entry(app, textvariable=url_input, width=100)
ent_OpenCatalog.insert(0, "http://open.isogeo.com/s/ad6451f1f9ca405ca6f78fabf46aeb10/Bue0ySfhmGOPw33jHMyaJtcOM4MY0")
ent_OpenCatalog.pack()
ent_OpenCatalog.focus_set()

# bouton
Button(app, text="Excelization!", command=lambda: app.destroy()).pack()

# initialisation de l'UI
app.mainloop()

##################### Excel sheet creation

##### Writing into an Excel file
book = xlwt.Workbook(encoding='utf8')
book.set_owner(str('Isogeo'))

# styles
style_header = xlwt.easyxf('pattern: pattern solid, fore_colour black;'
                           'font: colour white, bold True, height 220;'
                           'align: horiz center')
style_url = xlwt.easyxf(u'font: underline single')
style_wrap = xlwt.easyxf('align: wrap True')

# sheets
sheet_mds = book.add_sheet('Metadonnées', cell_overwrite_ok=True)

# headers
sheet_mds.write(0, 0, "Titre", style_header)
sheet_mds.write(0, 1, "Nom de la ressource", style_header)
sheet_mds.write(0, 2, "Emplacement", style_header)
sheet_mds.write(0, 3, "Mots-clés", style_header)
sheet_mds.write(0, 4, "Résumé", style_header)
sheet_mds.write(0, 5, "Thématiques INPIRES", style_header)
sheet_mds.write(0, 6, "Type", style_header)
sheet_mds.write(0, 7, "Format", style_header)
sheet_mds.write(0, 8, "SRS", style_header)
sheet_mds.write(0, 9, "Nombre d'objets", style_header)
sheet_mds.write(0, 10, "Géométrie", style_header)
sheet_mds.write(0, 11, "Propriétaire", style_header)
sheet_mds.write(0, 12, "Données - Création", style_header)
sheet_mds.write(0, 13, "Données - Modification", style_header)
sheet_mds.write(0, 14, "Métadonnées - Création", style_header)
sheet_mds.write(0, 15, "Métadonnées - Modification", style_header)
sheet_mds.write(0, 16, "Conformité INSPIRE", style_header)
sheet_mds.write(0, 17, "# contacts", style_header)
sheet_mds.write(0, 18, "Points de contacts", style_header)
sheet_mds.write(0, 19, "Points de contacts", style_header)
sheet_mds.write(0, 20, "Visualiser sur l'OpenCatalog", style_header)
sheet_mds.write(0, 21, "Editer sur Isogeo", style_header)

# columns width
sheet_mds.col(0).width = 50 * 100
sheet_mds.col(1).width = 40 * 256
sheet_mds.col(4).width = 75 * 256

##################### Calling Isogeo API

# get the OpenCatalog URL given
url_OpenCatalog = url_input.get()
if not url_OpenCatalog[-1] == '/':
    url_OpenCatalog = url_OpenCatalog + '/'
else:
    pass

# get the clean
url_base = url_OpenCatalog[0:url_OpenCatalog.index(url_OpenCatalog.rsplit('/')[6])]

# isoler l’identifiant du partage
share_id = url_OpenCatalog.rsplit('/')[4]
# isoler le token du partage
share_token = url_OpenCatalog.rsplit('/')[5]

# test if URL already contains some filters
if len(url_OpenCatalog.rsplit('/')) == 8:
    filters = url_OpenCatalog.rsplit('/')[7]
else:
    filters = ""
    pass

# setting the psubresources to include
includes = "conditions,contacts,coordinate-system,events,feature-attributes,keywords,limitations,links,specifications"

# écriture de la requête de recherche à l'API
search_req = Request("http://v1.api.isogeo.com/resources/search?ct={0}&s={1}&q={2}&_limit=100&_lang={3}&_offset={4}&_include={5}".format(share_token, share_id, filters, lang, start, includes))

# requête pour les caractéristiques du partage
share_req = Request('https://v1.api.isogeo.com/shares/{0}?token={1}'.format(share_id, share_token))

# envoi de la requête dans une boucle de test pour prévenir les erreurs
try:
    search_resp = urlopen(search_req)
    search_rez = json.load(search_resp)
    share_resp = urlopen(share_req)
    share_rez = json.load(share_resp)
except URLError, e:
    print(e)

if not search_rez:
    print("Request failed. Check your connection state.")
    exit()
else:
    pass

# share caracteristics
li_catalogs = share_rez.get("catalogs")

# tags
tags = search_rez.get('tags')
li_owners = [tags.get(tag) for tag in tags.keys() if tag.startswith('owner')]

# results
tot_results = search_rez.get('total')
print("Total :  ", tot_results)
metadatas = search_rez.get('results')
li_ids_md = [md.get('_id') for md in metadatas]

# handling Isogeo API limit
# reference: https://docs.google.com/document/d/11dayY1FH1NETn6mn9Pt2y3n8ywVUD0DoKbCi9ct9ZRo/edit#heading=h.bg6le8mcd07z
if tot_results > 100:
    # if API returned more than one page of results, let's get the rest!
    for idx in range(1, int(ceil(tot_results / 100)) + 1):
        start = idx * 100 + 1
        print(start)
        search_req = Request("https://v1.api.isogeo.com/resources/search?ct={0}&s={1}&q={2}&_limit=100&_lang={3}&_offset={4}&_include={5}".format(share_token, share_id, filters, lang, start, includes))
        try:
            search_resp = urlopen(search_req)
            search_rez = json.load(search_resp)
        except URLError, e:
            print(e)
        metadatas.extend(search_rez.get('results'))
else:
    pass

# passing parameters to the Excel function
md2wb(sheet_mds, 0, metadatas, li_catalogs)

# Sauvegarde du fichier Excel
dstamp = datetime.now()
book.save(r"output\isogeo2xls_{0}_{1}{2}{3}{4}{5}{6}.xls".format(share_rez.get("name"),
                                                                 dstamp.year,
                                                                 dstamp.month,
                                                                 dstamp.day,
                                                                 dstamp.hour,
                                                                 dstamp.minute,
                                                                 dstamp.second))

###############################################################################
###### Stand alone program ########
###################################

# if __name__ == '__main__':
#     """ standalone execution """
#     main()
