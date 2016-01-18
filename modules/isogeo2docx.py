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
import locale
from math import ceil
from os import listdir, path
from sys import exit
from Tkinter import Tk, StringVar
from ttk import Label, Button, Entry, Combobox    # widgets
from urllib2 import Request, urlopen, URLError

# 3rd party library
from dateutil.parser import parse as dtparse
from docxtpl import DocxTemplate

###############################################################################
########### Classes ###############
###################################

class Isogeo2docx(object):
    """
    docstring for Isogeo
    """
    def __init__(self, docx_template, search_results, url_base):
        """ Isogeo connection parameters

        docx_template -- application identifier
        search_results -- application
        url_base -- language asked for localized tags (INSPIRE themes)
        """
        super(Isogeo, self).__init__()
        self.id = client_id
        self.ct = client_secret


###############################################################################
########## Functions ##############
###################################


def md2docx(docx_template, offset, md, li_catalogs, url_base):
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
    inspire_valid = "Non"
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
            inspire_valid = "Oui"
            continue
        else:
            pass

    # formatting links to visualize on OpenCatalog and edit on APP
    link_visu = url_base + "m/" + md_id
    link_edit = "https://app.isogeo.com/resources/" + md_id

    # CONTACTS #
    contacts = md.get("contacts")
    # formatting contacts
    if len(contacts):
        contacts_cct = ["{5} {0} ({1})\n{2}\n{3}\n{4} ;\n\n".format(contact.get("contact").get("name"),
                                                                    contact.get("contact").get("organization"),
                                                                    contact.get("contact").get("email"),
                                                                    contact.get("contact").get("phone"),
                                                                    unicode(contact.get("contact").get("addressLine1"))
                                                                    + u", " + unicode(contact.get("contact").get("zipCode"))
                                                                    + u" " + unicode(contact.get("contact").get("city")),
                                                                    contact.get("role"))
                        for contact in contacts]
                        # for contact in contacts if contact.get("role") == "pointOfContact"]
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

    # CGUs AND lIMITATIONS #
    cgus = md.get("conditions")
    # formatting contacts
    if cgus:
        cgus_cct = ["{1} {0} ({2}) ;\n\n".format(cgu.get("description"),
                                                 cgu.get("license").get("name"),
                                                 cgu.get("license").get("link"))\
                    for cgu in cgus if cgu.get('license')]
    else:
        cgus_cct = ""

    limitations = md.get("limitations")
    # formatting contacts
    if limitations:
        limits_cct = ["Type : {0} - Restriction : {1} ;\n\n".format(lim.get("type"),
                                                                    lim.get("restriction"))\
                    for lim in limitations]
    else:
        limits_cct = ""

    # validity
    # for date manipulation: https://docs.python.org/2/library/datetime.html#strftime-strptime-behavior
    # could be independant from dateutil: datetime.datetime.strptime("2008-08-12T12:20:30.656234Z", "%Y-%m-%dT%H:%M:%S.Z")
    if md.get("validFrom"):
        valid_start = dtparse(md.get("validFrom")).strftime("%a %d %B %Y")
    else:
        valid_start = "NR"
    # end validity date
    if md.get("validTo"):
        valid_end = dtparse(md.get("validTo")).strftime("%a %d %B %Y")
    else:
        valid_end = "NR"
    # vailidty comment
    if md.get("validyComment"):
        valid_com = md.get("validyComment")
    else:
        valid_com = "NR"

    # METADATA #
    md_created = dtparse(md.get("_created")).strftime("%a %d %B %Y (%Hh%M)")
    md_updated = dtparse(md.get("_modified")).strftime("%a %d %B %Y (%Hh%M)")

    # FILLFULLING THE TEMPLATE #
    context = {
              'varTitle': md.get("title"),
              'varAbstract': md.get("abstract"),
              'varNameTech': md.get("name"),
              'varCollectContext': md.get("collectionContext"),
              'varCollectMethod': md.get("collectionMethod"),
              'varDataDtCrea': data_created.decode('latin1'),
              'varDataDtUpda': data_updated.decode('latin1'),
              'varDataDtPubl': data_published.decode('latin1'),
              'varValidityStart': valid_start.decode('latin1'),
              'varValidityEnd': valid_end.decode('latin1'),
              'validityComment': valid_com,
              'varFormat': format_version,
              'varGeometry': md.get("geometry"),
              'varObjectsCount': md.get("features"),
              'varKeywords': " ; ".join(li_motscles),
              'varKeywordsCount': len(li_motscles),
              'varType': md.get("type"),
              'varOwner': owner,
              'varScale': md.get("scale"),
              'varTopologyInfo': md.get("topologicalConsistency"),
              'varInspireTheme': " ; ".join(li_theminspire),
              'varInspireConformity': inspire_valid,
              'varInspireLimitation': " ; \n".join(limits_cct),
              'varCGUs': " ; \n".join(cgus_cct),
              'varContactsCount': len(contacts),
              'varContactsDetails': " ; \n".join(contacts_cct),
              'varSRS': srs,
              'varPath': localplace,
              'varFieldsCount': len(fields),
              'items': list(fields),
              'varMdDtCrea': md_created.decode('latin1'),
              'varMdDtUpda': md_updated.decode('latin1'),
              'varMdDtExp': datetime.now().strftime("%a %d %B %Y (%Hh%M)").decode('latin1'),
              'varViewOC': link_visu,
              'varEditAPP': link_edit,
              }

    # fillfull file
    try:
        docx_template.render(context)
    except Exception, e:
        print(u"Metadata error: check if there's any special character (<, <, &...) in different fields (attributes names and description...). Link: {0}".format(link_edit))
        print(e)

    # end of function
    return


def remove_accents(input_str, substitute=u""):
    """
    Clean string from special characters
    source: http://stackoverflow.com/a/5843560
    """
    return unicode(substitute).join(char for char in input_str if char.isalnum())

###############################################################################
######### Main program ############
###################################


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
ent_OpenCatalog.insert(0, "https://open.isogeo.com/s/ad6451f1f9ca405ca6f78fabf46aeb10/Bue0ySfhmGOPw33jHMyaJtcOM4MY0/q/keyword:inspire-theme:administrativeunits")
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




###############################################################################
###### Stand alone program ########
###################################

# if __name__ == '__main__':
#     """ standalone execution """
#     main()
