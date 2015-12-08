# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (print_function, unicode_literals)
#------------------------------------------------------------------------------
# Name:         OpenCatalog to Excel
# Purpose:      Get metadatas from an Isogeo OpenCatlog and store it into
#               an Excel workbook.
#
# Author:       Julien Moura (@geojulien) & Valentin Blanlot (@bablot)
#
# Python:       2.7.x
# Created:      14/08/2014
# Updated:      22/12/2015
#------------------------------------------------------------------------------

###############################################################################
########### Libraries #############
###################################

# Standard library
from datetime import datetime
import json
from math import ceil
from os import listdir, path
from sys import exit
from Tkinter import Tk, StringVar, OptionMenu
from ttk import Label, Button, Entry, Combobox    # widgets
from urllib2 import Request, urlopen, URLError

# 3rd party library
from docxtpl import DocxTemplate

###############################################################################
########## Functions ##############
###################################

def md2docx(docx_template, offset, md, li_catalogs):
    """
    parses Isogeo metadatas and replace docx template
    """
    # optional: print resource id (useful in debug mode)
    md_id = md.get("_id")
    print(md_id)

    # TAGS #
    # extracting & parsing tags
    tags = md.get("tags")
    li_motscles = []
    li_theminspire = []
    srs = ""
    owner = ""
    inspire_valid = 0
    format_lbl = ""
    fields = ["NR"]

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
        if tag.startswith('format'):
            format_lbl = tags.get(tag)
            continue
        else:
            pass
        # INSPIRE conformity
        if tag.startswith('conformity:inspire'):
            inspire_valid = 1
            continue
        else:
            pass

    # formatting links to visualize on OpenCatalog and edit on APP
    # link_visu = url_OpenCatalog + "/m/" + md_id
    link_edit = "https://app.isogeo.com/resources/" + md_id

    # CONTACTS #
    contacts = md.get("contacts")
    # formatting contacts
    if len(contacts):
        contacts_cct = ["{0} ({1}) ;\n".format(contact.get("contact").get("name"),
                                               contact.get("contact").get("email"))\
                        for contact in contacts if contact.get("role") == "pointOfContact"]
    else:
        contacts_cct = ""

    # ATTRIBUTES #
    # formatting feature attributes
    if md.get("type") == "vectorDataset" and md.get("feature-attributes"):
        fields = md.get("feature-attributes")
    else:
        fields = []
        pass

    # IDENTIFICATION #
    # format version
    if md.get("formatVersion"):
        format_version = u"{0} ({1} - {2})".format(format_lbl,
                                                   md.get("formatVersion"),
                                                   md.get("encoding"))
    else:
        format_version = format_lbl

    # path to the resource
    if md.get("path"):
        localplace = md.get("path").replace("&", "&amp;")
    else:
        localplace = 'NR'

    # HISTORY #
    # data events
    data_created = md.get("created")
    data_updated = md.get("modified")
    data_published = md.get("published")

    # METADATA #
    md_created = md.get("_created")
    md_updated = md.get("_modified")

    # FILLFULLING THE TEMPLATE #
    context = {
              'varTitle': md.get("title"),
              'varAbstract': md.get("abstract"),
              'varNameTech': md.get("name"),
              'varDataDtCrea': data_created,
              'varDataDtUpda': data_updated,
              'varDataDtPubl': data_published,
              'varFormat': format_version,
              'varGeometry': md.get("geometry"),
              'varObjectsCount': md.get("features"),
              'varKeywords': " ; ".join(li_motscles),
              'varType': md.get("type"),
              'varOwner': owner,
              'varScale': md.get("scale"),
              'varInspireTheme': " ; ".join(li_theminspire),
              'varInspireConformity': inspire_valid,
              'varContactsCount': len(contacts),
              'varContactsDetails': " ; \n".join(contacts_cct),
              'varSRS': srs,
              'varPath': localplace,
              'varFieldsCount': len(fields),
              'items': list(fields),
              'varMdDtCrea': md_created,
              'varMdDtUpda': md_updated,
              'varMdDtExp': datetime.now(),
              }

    # fillfull file
    try:
        docx_template.render(context)
        print('OK')
    except:
        print(u"Metadata error: check if there's any special character (<, <, &...) in different fields (attributes names and description...). Link: {0}".format(link_edit))

    # end of function
    return

###############################################################################
######### Main program ############
###################################

# list available templates
from os import getcwd
templates = [path.abspath(path.join(r'templates', tpl)) for tpl in listdir(r'templates') if path.splitext(tpl)[1].lower() == ".docx"]  # languages


##################### UI
app = Tk()
app.title('OpenCatalog ===> Word')

# variables
url_input = StringVar(app)
tpl_input = StringVar(app)
lang = "fr"
start = 0

# étiquette
lb_input_oc = Label(app, text="Coller l'URL d'un OpenCatalog").pack()

# champ pour l'URL
ent_OpenCatalog = Entry(app, textvariable=url_input, width=100)
ent_OpenCatalog.insert(0, "http://open.isogeo.com/s/ad6451f1f9ca405ca6f78fabf46aeb10/Bue0ySfhmGOPw33jHMyaJtcOM4MY0")
ent_OpenCatalog.pack()
ent_OpenCatalog.focus_set()

# pick a template
lb_input_tpl = Label(app, text="Choisir un template").pack()
droplist = Combobox(app,
                textvariable=tpl_input,
                values=templates,
                width=100)
droplist.pack()

# bouton
Button(app, text="Wordification !", command=lambda: app.destroy()).pack()

# initialisation de l'UI
app.mainloop()

##################### Calling Isogeo API

# get the OpenCatalog URL
url_OpenCatalog = url_input.get()

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

# écriture de la requête de recherche à l'API
search_req = Request("http://v1.api.isogeo.com/resources/search?ct={0}&s={1}&q={2}&_limit=100&_lang={3}&_offset={4}&_include=contacts,links,feature-attributes".format(share_token,
                                                                                                                                                                 share_id,
                                                                                                                                                                 filters,
                                                                                                                                                                 lang,
                                                                                                                                                                 start))

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
li_catalogs = share_rez.get("catalogs")

# tags
tags = search_rez.get('tags')
li_owners = [tags.get(tag) for tag in tags.keys() if tag.startswith('owner')]

# results
tot_results = search_rez.get('total')
metadatas = search_rez.get('results')
li_ids_md = [md.get('_id') for md in metadatas]

# handling Isogeo API limit
# reference: https://docs.google.com/document/d/11dayY1FH1NETn6mn9Pt2y3n8ywVUD0DoKbCi9ct9ZRo/edit#heading=h.bg6le8mcd07z
if tot_results > 100:
    # if API returned more than one page of results, let's get the rest!
    for idx in range(1, int(ceil(tot_results / 100)) + 1):
        start = idx * 100 + 1
        print(start)
        search_req = Request("http://v1.api.isogeo.com/resources/search?ct={0}&s={1}&q={2}&_limit=100&_lang={3}&_offset={4}&_include=contacts,links,feature-attributes".format(share_token,
                                                                                                                                                                               share_id,
                                                                                                                                                                               filters,
                                                                                                                                                                               lang,
                                                                                                                                                                               start))
        try:
            search_resp = urlopen(search_req)
            search_rez = json.load(search_resp)
        except URLError, e:
            print(e)
        metadatas.extend(search_rez.get('results'))
else:
    pass

print(tpl_input.get())

# setting Word
for md in metadatas:
    docx_tpl = DocxTemplate(path.realpath(tpl_input.get()))
    dstamp = datetime.now()
    md2docx(docx_tpl, 0, md, li_catalogs)  # passing parameters to the Word generator
    docx_tpl.save(r"output\{0}_{7}_{1}{2}{3}{4}{5}{6}.docx".format(share_rez.get("name"),
                                                                   dstamp.year,
                                                                   dstamp.month,
                                                                   dstamp.day,
                                                                   dstamp.hour,
                                                                   dstamp.minute,
                                                                   dstamp.second,
                                                                   md.get("_id")[:8]))

###############################################################################
###### Stand alone program ########
###################################

# if __name__ == '__main__':
#     """ standalone execution """
#     main()
