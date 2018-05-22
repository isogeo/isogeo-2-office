# -*- coding: UTF-8 -*-

# ------------------------------------------------------------------------------
# Name:         Isogeo to Microsoft Word 2010
# Purpose:      Get metadatas from an Isogeo share and store it into
#               a Word document for each metadata. It's one of the submodules
#               of isogeo2office (https://github.com/isogeo/isogeo-2-office).
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
from itertools import zip_longest
import logging
import re
from xml.sax.saxutils import escape

# 3rd party library
import arrow
from isogeo_pysdk import IsogeoTranslator


# ##############################################################################
# ############ Globals ############
# #################################

logger = logging.getLogger("isogeo2office")  # LOG

# ##############################################################################
# ########## Classes ###############
# ##################################


class IsogeoFormatter(object):
    """IsogeoFormatter class."""

    def __init__(self, lang="FR",
                 output_type="Excel",
                 default_values=("NR", "1970-01-01T00:00:00+00:00")):
        """Metadata formatter to avoid repeat oeprations on metadata.

        :param str lang: selected language
        :param str output_type: name of output type to format for
        :param tuple default_values: values used to replace missing values.
         2 values structure:
            (
            str_for_missing_strings_and_integers,
            str_for_missing_dates
            )
        """
        # locale
        self.lang = lang
        if lang.lower() == "fr":
            self.dates_fmt = "DD/MM/YYYY"
            self.locale_fmt = "fr_FR"
        else:
            self.dates_fmt = "YYYY/MM/DD"
            self.locale_fmt = "uk_UK"

        # store params and imports as attributes
        self.output_type = output_type
        self.defs = default_values
        self.tr = IsogeoTranslator(lang).tr

    # ------------ Metadata sections formatter --------------------------------
    def conditions(self, md_cgus: dict):
        """Render input metadata CGUs as a new list.

        :param dict md_cgus: input dictionary extracted from an Isogeo metadata
        """
        cgus_out = []
        for c_in in md_cgus:
            cgu = {}
            # ensure other fields
            cgu["description"] = self.clean_xml(c_in.get("description", ""))
            if "license" in c_in.keys():
                cgu["name"] = self.clean_xml(c_in.get("license").get("name", "NR"))
                cgu["link"] = c_in.get("license").get("link", "")
                cgu["content"] = self.clean_xml(c_in.get("license").get("content", ""))
            else:
                cgu["name"] = self.tr("conditions", "noLicense")

            # store into the final list
            cgus_out.append("{} {}. {} {}".format(cgu.get("name"),
                                                  cgu.get("description", ""),
                                                  cgu.get("content", ""),
                                                  cgu.get("link", "")))
        # return formatted result
        return cgus_out

    def limitations(self, md_limitations: dict):
        """Render input metadata limitations as a new list.

        :param dict md_limitations: input dictionary extracted from an Isogeo metadata
        """
        lims_out = []
        for l_in in md_limitations:
            limitation = {}
            # ensure other fields
            limitation["description"] = self.clean_xml(l_in.get("description", ""))
            limitation["type"] = self.tr("limitations", l_in.get("type"))
            # legal type
            if l_in.get("type") == "legal":
                limitation["restriction"] = self.tr("restrictions",
                                                    l_in.get("restriction"))
            else:
                pass
            # INSPIRE precision
            if "directive" in l_in.keys():
                limitation["inspire"] = self.clean_xml(l_in.get("directive")
                                                           .get("name"))
                limitation["content"] = self.clean_xml(l_in.get("directive")
                                                           .get("description"))
            else:
                pass

            # store into the final list
            lims_out.append("{} {}. {} {} {}".format(limitation.get("type"),
                                                     limitation.get("description", ""),
                                                     limitation.get("restriction", ""),
                                                     limitation.get("content", ""),
                                                     limitation.get("inspire", "")))
        # return formatted result
        return lims_out

    def specifications(self, md_specifications: dict):
        """Render input metadata specifications as a new list.

        :param dict md_specifications: input dictionary extracted from an Isogeo metadata
        """
        specs_out = []
        for s_in in md_specifications:
            spec = {}
            # translate specification conformity
            if s_in.get("conformant"):
                spec["conformity"] = self.tr("quality", "isConform")
            else:
                spec["conformity"] = self.tr("quality", "isNotConform")
            # ensure other fields
            spec["name"] = s_in.get("specification").get("name")
            spec["link"] = s_in.get("specification").get("link")
            # make data human readable
            spec_date = arrow.get(s_in.get("specification")
                                      .get("published")[:19])
            spec_date = "{0}".format(spec_date.format(self.dates_fmt,
                                                      self.locale_fmt))
            spec["date"] = spec_date
            # store into the final list
            specs_out.append("{} {} {} - {}"
                             .format(spec.get("name"),
                                     spec.get("date"),
                                     spec.get("link"),
                                     spec.get("conformity")))

        # return formatted result
        return specs_out

    # ------------ Prevent encoding errors ------------------------------------
    def remove_accents(self, input_str, substitute=u""):
        """Clean string from special characters.

        source: http://stackoverflow.com/a/5843560
        """
        return unicode(substitute).join(char for char in input_str
                                        if char.isalnum())

    def clean_xml(self, invalid_xml, mode="soft", substitute="_"):
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
        pairs = zip_longest(*iters, fillvalue='')  # iterate 2 items at a time

        # get the clean version
        clean_version = ''.join(escape(text) + tag for text, tag in pairs)
        if mode == "strict":
            clean_version = re.sub(r"<.*?>", substitute, clean_version)
        else:
            pass
        return clean_version


# ###############################################################################
# ###### Stand alone program ########
# ###################################
if __name__ == '__main__':
    """Try me"""
    formatter = IsogeoFormatter()
