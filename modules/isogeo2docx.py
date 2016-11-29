# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)
# ------------------------------------------------------------------------------
# Name:         Isogeo to Microsoft Word 2010
# Purpose:      Get metadatas from an Isogeo share and store it into
#               a Word document for each metadata. It's one of the submodules
#               of isogeo2office (https://bitbucket.org/isogeo/isogeo-2-office).
#
# Author:       Julien Moura (@geojulien)
#
# Python:       2.7.x
# Created:      14/08/2014
# Updated:      28/01/2016
# ------------------------------------------------------------------------------

# ##############################################################################
# ########## Libraries #############
# ##################################

# Standard library
from datetime import datetime
from itertools import izip_longest
import logging
import re
from xml.sax.saxutils import escape  # '<' -> '&lt;'

# 3rd party library
import arrow
from docxtpl import DocxTemplate, etree
from isogeo_pysdk import Isogeo

# custom
from isogeo_api_strings import dict_md_fields_fr

# ##############################################################################
# ########## Classes ###############
# ##################################


class Isogeo2docx(object):
    """IsogeoToDocx class."""

    def __init__(self, lang="FR",
                 default_values=("NR", "1970-01-01T00:00:00+00:00")):
        """Common variables for Word processing.

        default_values (optional) -- values used to replace missing values.
        Must be a tuple with 2 values structure:
            (
            str_for_missing_strings_and_integers,
            str_for_missing_dates
            )
        """
        super(Isogeo2docx, self).__init__()

        # ------------ VARIABLES ---------------------
        # test variables
        if type(default_values) != tuple:
            raise TypeError(self.__init__.__doc__)
        else:
            pass
        if len(default_values) != 2:
            raise ValueError(self.__init__.__doc__)
        else:
            pass

        # set variables
        self.default_values = default_values

        # LOCALE
        if lang == "FR":
            self.dates_fmt = "DD/MM/YYYY"
            self.locale_fmt = "fr_FR"
        else:
            self.dates_fmt = "YYYY/MM/DD"
            self.locale_fmt = "uk_UK"

    def md2docx(self, docx_template, md, url_base):
        """Parse Isogeo metadatas and replace docx template."""
        # optional: print resource id (useful in debug mode)
        md_id = md.get("_id")

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
                owner_id = tag[6:]
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
                format_lbl = tags.get(tag, self.missing_values())
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
        link_visu = url_base + "/m/" + md_id
        link_edit = "https://app.isogeo.com/groups/{}/resources/{}".format(owner_id, md_id)

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
            # prevent invalid character for XML formatting
            for f in fields:
                for i in f.keys():
                    t = f.get(i)
                    if type(t) in (str, unicode):
                        f[i] = self.clean_xml(t)
                    else:
                        pass
        else:
            fields = []
            pass

        # EVENTS #
        # formatting feature attributes
        events = md.get("events", "")
        # print(events)
        for e in events:
            # pop creation events (already in the export document)
            if e.get("kind") == "creation":
                events.remove(e)
                continue
            else:
                pass
            # prevent invalid character for XML formatting in description
            e["description"] = self.clean_xml(e.get("description", ""))
            # make data human readable
            evt_date = arrow.get(e.get("date")[:19])
            evt_date = "{0} ({1})".format(evt_date.format(self.dates_fmt,
                                                         self.locale_fmt),
                                          evt_date.humanize(locale=self.locale_fmt))
            e["date"] = evt_date
            # translate event kind
            e["kind"] = dict_md_fields_fr.get("events")\
                                         .get(e.get("kind"))




        

        # IDENTIFICATION #
        # format version
        # if md.get("formatVersion"):
        #     format_version = u"{0} ({1} - {2})".format(format_lbl,
        #                                                md.get("formatVersion", "NR"),
        #                                                md.get("encoding", "NR"))
        # else:
        #     format_version = format_lbl

        format_version = u"{0} ({1} - {2})".format(format_lbl,
                                                   md.get("formatVersion",
                                                          self.missing_values()
                                                          ),
                                                   md.get("encoding",
                                                          self.missing_values()
                                                          )
                                                   )

        # path to the resource
        localplace = md.get("path", self.missing_values()).replace("&", "&amp;")
        # if md.get("path"):
        #     localplace = md.get("path").replace("&", "&amp;")
        # else:
        #     localplace = 'NR'

        # HISTORY #
        # data events

        if md.get("created"):
            data_created = arrow.get(md.get("created")[:19])
            data_created = "{0} ({1})".format(data_created.format(self.dates_fmt,
                                                                  self.locale_fmt),
                                              data_created.humanize(locale=self.locale_fmt))
            # data_created = arrow.get(md.get("created", self.missing_values(1))).format("dddd D MMMM YYYY")
        else:
            data_created = "NR"
        if md.get("modified"):
            data_updated = arrow.get(md.get("_created")[:19])
            data_updated = "{0} ({1})".format(data_updated.format(self.dates_fmt,
                                                                  self.locale_fmt),
                                            data_updated.humanize(locale=self.locale_fmt))
        else:
            data_updated = "NR"
        if md.get("published"):
            data_published = arrow.get(md.get("_created")[:19])
            data_published = "{0} ({1})".format(data_published.format(self.dates_fmt,
                                                              self.locale_fmt),
                                                data_published.humanize(locale=self.locale_fmt))
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
            valid_start = arrow.get(md.get("validFrom"))
            valid_start = "{0}".format(valid_start.format(self.dates_fmt, self.locale_fmt))
        else:
            valid_start = "NR"
        # end validity date
        if md.get("validTo"):
            valid_end = arrow.get(md.get("validTo"))
            valid_end = "{0}".format(valid_end.format(self.dates_fmt, self.locale_fmt))
        else:
            valid_end = "NR"
        # validity comment
        valid_com = md.get("validityComment", self.missing_values())

        # METADATA #
        md_created = arrow.get(md.get("_created")[:19])
        md_created = "{0} ({1})".format(md_created.format(self.dates_fmt,
                                                          self.locale_fmt),
                                        md_created.humanize(locale=self.locale_fmt))
        md_updated = arrow.get(md.get("_modified")[:19])
        md_updated = "{0} ({1})".format(md_updated.format(self.dates_fmt,
                                                          self.locale_fmt),
                                        md_updated.humanize(locale=self.locale_fmt))

        # FILLFULLING THE TEMPLATE #
        context = {
                  'varTitle': self.clean_xml(md.get("title", self.missing_values())),
                  'varAbstract': self.clean_xml(md.get("abstract", self.missing_values())),
                  'varNameTech': md.get("name", self.missing_values()),
                  'varCollectContext': self.clean_xml(md.get("collectionContext", self.missing_values())),
                  'varCollectMethod': self.clean_xml(md.get("collectionMethod", self.missing_values())),
                  'varDataDtCrea': data_created.decode('latin1'),
                  'varDataDtUpda': data_updated.decode('latin1'),
                  'varDataDtPubl': data_published.decode('latin1'),
                  'varValidityStart': valid_start.decode('latin1'),
                  'varValidityEnd': valid_end.decode('latin1'),
                  'validityComment': self.clean_xml(valid_com),
                  'varFormat': format_version,
                  'varGeometry': md.get("geometry", self.missing_values()),
                  'varObjectsCount': md.get("features", self.missing_values()),
                  'varKeywords': " ; ".join(li_motscles),
                  'varKeywordsCount': len(li_motscles),
                  'varType': md.get("type", self.missing_values()),
                  'varOwner': owner,
                  'varScale': md.get("scale", self.missing_values()),
                  'varTopologyInfo': self.clean_xml(md.get("topologicalConsistency", self.missing_values())),
                  'varInspireTheme': " ; ".join(li_theminspire),
                  'varInspireConformity': inspire_valid,
                  'varInspireLimitation': " ; \n".join(limits_cct),
                  'varCGUs': self.clean_xml(" ; \n".join(cgus_cct)),
                  'varContactsCount': len(contacts),
                  'varContactsDetails': " ; \n".join(contacts_cct),
                  'varSRS': srs,
                  'varPath': localplace,
                  'varFieldsCount': len(fields),
                  'fields': list(fields),
                  'varEventsCount': len(events),
                  'events': list(events),
                  'varMdDtCrea': md_created.decode('latin1'),
                  'varMdDtUpda': md_updated.decode('latin1'),
                  'varMdDtExp': datetime.now().strftime("%a %d %B %Y (%Hh%M)").decode('latin1'),
                  'varViewOC': link_visu,
                  'varEditAPP': link_edit,
                  }

        # fillfull file
        try:
            docx_template.render(context)
            logging.info("Vector metadata stored: {} ({})".format(md.get("name"),
                                                                  md.get("_id")))
        except etree.XMLSyntaxError as e:
            logging.error("Invalid character in XML: {}. "
                          "Any special character (<, <, &...)? Check: {}".format(e, link_edit))
        except (UnicodeEncodeError, UnicodeDecodeError) as e:
            logging.error("Encoding error: {}. "
                          "Any special character (<, <, &...)? Check: {}".format(e, link_edit))
        except Exception as e:
            logging.error("Unexpected error: {}. Check: {}".format(e, link_edit))

        # end of function
        return

    # ------------ UTILITIES ---------------------
    def missing_values(self, idx_type=0):
        """Return default values defined in the class as a tuple.

        idx_type (optional) -- index of the value type requested:

            1: for strings and integers
            2: for dates and datetimes
        """
        rpl_value = self.default_values[idx_type]
        # end of method
        return rpl_value

    def remove_accents(self, input_str, substitute=u""):
        """Clean string from special characters.

        source: http://stackoverflow.com/a/5843560
        """
        return unicode(substitute).join(char for char in input_str if char.isalnum())

    def clean_xml(self, invalid_xml):
        """Clean string of XML invalid characters.

        source: http://stackoverflow.com/a/13322581/2556577
        """
        # assumptions:
        #   doc = *( start_tag / end_tag / text )
        #   start_tag = '<' name *attr [ '/' ] '>'
        #   end_tag = '<' '/' name '>'
        ws = r'[ \t\r\n]*'  # allow ws between any token
        name = '[a-zA-Z]+'  # note: expand if necessary but the stricter the better
        attr = '{name} {ws} = {ws} "[^"]*"'  # note: fragile against missing '"'; no "'"
        start_tag = '< {ws} {name} {ws} (?:{attr} {ws})* /? {ws} >'
        end_tag = '{ws}'.join(['<', '/', '{name}', '>'])
        tag = '{start_tag} | {end_tag}'

        assert '{{' not in tag
        while '{' in tag:   # unwrap definitions
            tag = tag.format(**vars())

        tag_regex = re.compile('(%s)' % tag, flags=re.VERBOSE)

        # escape &, <, > in the text
        iters = [iter(tag_regex.split(invalid_xml))] * 2
        pairs = izip_longest(*iters, fillvalue='')  # iterate 2 items at a time

        # get the clean version
        return ''.join(escape(text) + tag for text, tag in pairs)


# ###############################################################################
# ###### Stand alone program ########
# ###################################

if __name__ == '__main__':
    """ Standalone execution and tests
    """
    # ------------ Specific imports ---------------------
    from ConfigParser import SafeConfigParser   # to manage options.ini
    from os import path

    # ------------ Settings from ini file ----------------
    if not path.isfile(path.realpath(r"..\settings.ini")):
        logging.error("To execute this script as standalone,"
                      " you need to store your Isogeo application settings"
                      " in a isogeo_params.ini file. You can use the template"
                      " to set your own.")
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
    search_results = isogeo.search(token,
                                   sub_resources=isogeo.sub_resources_available)

    # ------------ REAL START ----------------------------
    url_oc = "http://open.isogeo.com/s/c502e8f7c9da4c3aacdf3d905672d54c/Q4SvPfiIIslbdwkbWRFJLk7XWo4G0/"
    toDocx = Isogeo2docx()

    for md in search_results.get("results")[:11]:
        tpl = DocxTemplate(path.realpath(r"..\templates\template_Isogeo.docx"))
        toDocx.md2docx(tpl, md, url_oc)
        dstamp = datetime.now()
        if not md.get('name'):
            md_name = "NR"
        elif '.' in md.get('name'):
            md_name = md.get("name").split(".")[1]
        else:
            md_name = md.get("name")
        tpl.save(r"..\output\{0}_{8}_{7}_{1}{2}{3}{4}{5}{6}.docx".format("TestDemoDev",
                                                                         dstamp.year,
                                                                         dstamp.month,
                                                                         dstamp.day,
                                                                         dstamp.hour,
                                                                         dstamp.minute,
                                                                         dstamp.second,
                                                                         md.get("_id")[:5],
                                                                         md_name))
        del tpl
