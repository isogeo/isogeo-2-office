# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import unicode_literals
#------------------------------------------------------------------------------
# Name:         OpenCatalog to Excel
# Purpose:      Get metadata from an Isogeo OpenCatlog and store it into
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
import json
from Tkinter import Tk, StringVar
from ttk import Label, Button, Entry # widgets
from urllib2 import Request, urlopen, URLError

import os
import platform

# 3rd party library
import xlwt

###############################################################################
######### Main program ############
###################################


##################### UI
app = Tk()
app.title('OpenCatalog ===> Excel')

# variables
url_input = StringVar(app)

# étiquette
lb_invite = Label(app, text="Colle ici ton OpenCatalog")
lb_invite.pack()

# champ pour l'URL
ent_OpenCatalog = Entry(app, textvariable=url_input, width=100)
ent_OpenCatalog.pack()

# bouton
Button(app, text="Excelization!", command=lambda:app.destroy()).pack()

# initialisation de l'UI
app.mainloop()

##################### Excelization

# copier/coller l’url de l’OpenCatalog créé
url_OpenCatalog = url_input.get()

# isoler l’identifiant du partage
share_id = url_OpenCatalog.rsplit('/')[4]
# isoler le token du partage
share_token = url_OpenCatalog.rsplit('/')[5]


#### Exemple sur un OpenCatalog
# écriture de la requête à l'API
search_req = Request('http://api.isogeo.com/v1.0/resources/search?ct={0}&s={1}'.format(share_token, share_id))

# envoi de la requête dans une boucle de test pour prévenir les erreurs
try:
    search_resp = urlopen(search_req)
    search_rez = json.load(search_resp)
except URLError, e:
    print e

# tags
tags = search_rez.get('tags')
li_owners = [tags.get(tag) for tag in tags.keys() if tag.startswith('owner')]

# results
metadatas = search_rez.get('results')
li_ids_md = [md.get('_id') for md in metadatas]


##### Writing into an Excel file
book = xlwt.Workbook(encoding='utf8')
book.set_owner(str('Isogeo - ') + str(','.join(li_owners)))

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
sheet_mds.write(0, 10, "Visualiser sur l'OpenCatalog", style_header)
sheet_mds.write(0, 11, "Editer sur Isogeo", style_header)

# writing contents
xls_line = 0

for md in metadatas:
    # incrémente le numéro de ligne
    xls_line += 1
    # extraction des mots-clés et thématiques
    tags = md.get("tags")
    li_motscles = [tags.get(tag) for tag in tags.keys() if tag.startswith('keyword:isogeo')]
    li_theminspire = [tags.get(tag) for tag in tags.keys() if tag.startswith('keyword:inspire-theme')]

    # formatage des liens pour visualiser et éditer
    link_visu = 'HYPERLINK("{0}"; "{1}")'.format(url_OpenCatalog + "/m/" + md.get('_id'),
                                            "Visualiser")
    link_edit = 'HYPERLINK("{0}"; "{1}")'.format("https://app.isogeo.com/resources/" + md.get('_id'),
                                            "Editer")
    # écriture des informations dans chaque colonne correspondante
    sheet_mds.write(xls_line, 0, md.get("title"))
    sheet_mds.write(xls_line, 1, md.get("name"))
    sheet_mds.write(xls_line, 2, md.get("path"))
    sheet_mds.write(xls_line, 3, " ; ".join(li_motscles))
    sheet_mds.write(xls_line, 4, md.get("abstract"))
    sheet_mds.write(xls_line, 5, " ; ".join(li_theminspire))
    sheet_mds.write(xls_line, 6, md.get("type"))
    sheet_mds.write(xls_line, 7, md.get("format"))
    sheet_mds.write(xls_line, 9, md.get("features"))
    sheet_mds.write(xls_line, 10, xlwt.Formula(link_visu), style_url)
    sheet_mds.write(xls_line, 11, xlwt.Formula(link_edit), style_url)

# Sauvegarde du fichier Excel
userhome = os.path.expanduser('~')
desktop = userhome + '/Desktop/'
book.save(desktop + r"OpenCatalog2excel.xls")

###############################################################################
###### Stand alone program ########
###################################

# if __name__ == '__main__':
#     """ standalone execution """
#     main()
