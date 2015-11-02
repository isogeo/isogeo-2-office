# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import unicode_literals
#------------------------------------------------------------------------------
# Name:         OpenCatalog to Excel
# Purpose:      Get metadatas from an Isogeo OpenCatlog and store it into
#               an Excel workbook.
#
# Author:       Julien Moura (@geojulien) & Valentin Blanlot (@bablot)
#
# Python:       2.7.x
# Created:      14/08/2014
# Updated:      22/12/2014
#------------------------------------------------------------------------------

###############################################################################
########### Libraries #############
###################################

# Standard library
from datetime import datetime
import json
from math import ceil
import os
from sys import exit
from Tkinter import Tk, StringVar
from ttk import Label, Button, Entry    # widgets
from urllib2 import Request, urlopen, URLError

# 3rd party library
import xlwt

###############################################################################
########## Functions ##############
###################################

def md2wb(wbsheet, offset, li_mds):
    """
    to describe
    """
    for md in li_mds:
        # incrémente le numéro de ligne
        offset += 1
        # extraction des mots-clés et thématiques
        tags = md.get("tags")
        li_motscles = []
        li_theminspire = []
        srs = ""
        owner = ""
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

        # formatage des liens pour visualiser et éditer
        link_visu = 'HYPERLINK("{0}"; "{1}")'.format(url_OpenCatalog + "/m/" + md.get('_id'),
                                                     "Visualiser")
        link_edit = 'HYPERLINK("{0}"; "{1}")'.format("https://app.isogeo.com/resources/" + md.get('_id'),
                                                     "Editer")
        # format version

        if md.get("formatVersion"):
            format_version = u"{0} ({1} - {2})".format(md.get("format"),
                                                       md.get("formatVersion"),
                                                       md.get("encoding"))
        else:
            format_version = u""

        # écriture des informations dans chaque colonne correspondante
        wbsheet.write(offset, 0, md.get("title"))
        wbsheet.write(offset, 1, md.get("name"))
        wbsheet.write(offset, 2, md.get("path"))
        wbsheet.write(offset, 3, " ; ".join(li_motscles))
        wbsheet.write(offset, 4, md.get("abstract"))
        wbsheet.write(offset, 5, " ; ".join(li_theminspire))
        wbsheet.write(offset, 6, md.get("type"))
        wbsheet.write(offset, 7, format_version)
        wbsheet.write(offset, 8, srs)
        wbsheet.write(offset, 9, md.get("features"))
        wbsheet.write(offset, 10, md.get("geometry"))
        wbsheet.write(offset, 11, owner)
        wbsheet.write(offset, 12, md.get("created"))
        wbsheet.write(offset, 13, md.get("modified"))
        wbsheet.write(offset, 14, md.get("_created"))
        wbsheet.write(offset, 15, md.get("_modified"))
        # wbsheet.write(offset, 16, md.get("_creator"))
        wbsheet.write(offset, 20, xlwt.Formula(link_visu), style_url)
        wbsheet.write(offset, 21, xlwt.Formula(link_edit), style_url)

    print(md.keys())
    print(md.get("links"))
    # end of function
    return

###############################################################################
######### Main program ############
###################################

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
sheet_mds.write(0, 20, "Visualiser sur l'OpenCatalog", style_header)
sheet_mds.write(0, 21, "Editer sur Isogeo", style_header)


########################

# copier/coller l’url de l’OpenCatalog créé
url_OpenCatalog = url_input.get()

# isoler l’identifiant du partage
share_id = url_OpenCatalog.rsplit('/')[4]
# isoler le token du partage
share_token = url_OpenCatalog.rsplit('/')[5]


#### Exemple sur un OpenCatalog
# écriture de la requête de recherche à l'API
search_req = Request('http://v1.api.isogeo.com/resources/search?ct={0}&s={1}&_limit=100&_lang={2}&_offset={3}'.format(share_token, share_id, lang, start))

# requête pour les caractéristiques du partage
share_req = Request('http://v1.api.isogeo.com/shares/{0}?token={1}'.format(share_id, share_token))

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
print(share_rez.keys())

# tags
tags = search_rez.get('tags')
li_owners = [tags.get(tag) for tag in tags.keys() if tag.startswith('owner')]

# results
tot_results = search_rez.get('total')
print(tot_results)
metadatas = search_rez.get('results')
li_ids_md = [md.get('_id') for md in metadatas]

# respecting Isogeo API limit
# reference: https://docs.google.com/document/d/11dayY1FH1NETn6mn9Pt2y3n8ywVUD0DoKbCi9ct9ZRo/edit#heading=h.bg6le8mcd07z
if tot_results > 100:
    # if API returned more than one page of results, let's get the rest!
    for idx in range(1, int(ceil(tot_results / 100)) + 1):
        start = idx * 100 + 1
        print(start)
        search_req = Request('http://v1.api.isogeo.com/resources/search?ct={0}&s={1}&_limit=100&_lang={2}&_offset={3}'.format(share_token, share_id, lang, start))
        try:
            search_resp = urlopen(search_req)
            search_rez = json.load(search_resp)
        except URLError, e:
            print(e)
        metadatas.extend(search_rez.get('results'))
else:
    pass

    # metalist_input = [metadatas[i:i + 100] for i in range(0, len(metadatas), 100)]
    # for sublist in metalist_input:
    #     md2wb(sheet_mds, 0, metadatas)

#
md2wb(sheet_mds, 0, metadatas)

# Sauvegarde du fichier Excel
userhome = os.path.expanduser('~')
desktop = userhome + '/Desktop/'
dstamp = datetime.now()
book.save(desktop + r"OpenCatalog2excel_{0}{1}{2}{3}{4}{5}.xls".format(dstamp.year,
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
