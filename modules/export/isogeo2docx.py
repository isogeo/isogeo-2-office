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
from xml.sax.saxutils import escape  # '<' -> '&lt;'

# 3rd party library
import arrow
from docxtpl import DocxTemplate, etree, InlineImage
from isogeo_pysdk import Isogeo, Metadata
from isogeo_pysdk import IsogeoTranslator

# custom submodules
from .formatter import IsogeoFormatter
from modules.utils import isogeo2office_utils

# ##############################################################################
# ############ Globals ############
# #################################

logger = logging.getLogger("isogeo2office")
utils = isogeo2office_utils()

# ##############################################################################
# ########## Classes ###############
# ##################################


class Isogeo2docx(object):
    """IsogeoToDocx class."""

    def __init__(self, lang="FR", default_values=("NR", "1970-01-01T00:00:00+00:00")):
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
        if lang.lower() == "fr":
            self.dates_fmt = "DD/MM/YYYY"
            self.locale_fmt = "fr_FR"
        else:
            self.dates_fmt = "YYYY/MM/DD"
            self.locale_fmt = "uk_UK"

        # TRANSLATIONS
        self.isogeo_tr = IsogeoTranslator(lang).tr

        # FORMATTER
        self.fmt = IsogeoFormatter(output_type="Word")

    def md2docx(self, docx_template, md: Metadata, url_base: str):
        """Parse Isogeo metadatas and replace docx template.
        
        
        """
        # optional: print resource id (useful in debug mode)
        md_id = md._id

        # TAGS #
        # extracting & parsing tags
        tags = md.tags
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
            if tag.startswith("keyword:isogeo"):
                li_motscles.append(tags.get(tag))
                continue
            else:
                pass
            # INSPIRE themes
            if tag.startswith("keyword:inspire-theme"):
                li_theminspire.append(tags.get(tag))
                continue
            else:
                pass
            # workgroup which owns the metadata
            if tag.startswith("owner"):
                owner = tags.get(tag)
                owner_id = tag[6:]
                continue
            else:
                pass
            # coordinate system
            if tag.startswith("coordinate-system"):
                srs = tags.get(tag)
                continue
            else:
                pass
            # format pretty print
            if tag.startswith("format"):
                format_lbl = tags.get(tag, self.missing_values())
                continue
            else:
                pass
            # INSPIRE conformity
            if tag.startswith("conformity:inspire"):
                inspire_valid = "Oui"
                continue
            else:
                pass

        # formatting links to visualize on OpenCatalog and edit on APP
        link_visu = url_base + "/m/" + md_id
        link_edit = "https://app.isogeo.com/groups/{}/resources/{}".format(
            owner_id, md_id
        )

        link_edit = "https://app.isogeo.com/groups/{}/resources/{}".format(
            owner_id, md_id
        )

        # ---- CONTACTS # ----------------------------------------------------
        if md.contacts:
            contacts_out = []
            # formatting contacts
            for ct_in in md.contacts:
                ct = {}
                # translate contact role
                ct["role"] = self.isogeo_tr("roles", ct_in.get("role"))
                # ensure other contacts fields
                ct["name"] = ct_in.get("contact").get("name", "NR")
                ct["organization"] = ct_in.get("contact").get("organization", "")
                ct["email"] = ct_in.get("contact").get("email", "")
                ct["phone"] = ct_in.get("contact").get("phone", "")
                ct["fax"] = ct_in.get("contact").get("fax", "")
                ct["addressLine1"] = ct_in.get("contact").get("addressLine1", "")
                ct["addressLine2"] = ct_in.get("contact").get("addressLine2", "")
                ct["zipCode"] = ct_in.get("contact").get("zipCode", "")
                ct["city"] = ct_in.get("contact").get("city", "")
                ct["countryCode"] = ct_in.get("contact").get("countryCode", "")
                # store into the final list
                contacts_out.append(ct)

        # ---- ATTRIBUTES --------------------------------------------------
        if md.type == "vectorDataset" and isinstance(md.featureAttributes, list):
            fields_out = []
            for f_in in md.featureAttributes:
                field = {}
                # ensure other fields
                field["name"] = self.clean_xml(f_in.get("name", ""))
                field["alias"] = self.clean_xml(f_in.get("alias", ""))
                field["description"] = self.clean_xml(f_in.get("description", ""))
                field["dataType"] = f_in.get("dataType", "")
                field["language"] = f_in.get("language", "")
                # store into the final list
                fields_out.append(field)

        # ---- EVENTS ------------------------------------------------------
        if md.events:
            for e in md.events:
                # pop creation events (already in the export document)
                if e.get("kind") == "creation":
                    md.events.remove(e)
                    continue
                else:
                    pass
                # prevent invalid character for XML formatting in description
                e["description"] = self.clean_xml(
                    e.get("description", " "), mode="strict", substitute=""
                )
                # make data human readable
                evt_date = arrow.get(e.get("date")[:19])
                evt_date = "{0} ({1})".format(
                    evt_date.format(self.dates_fmt, self.locale_fmt),
                    evt_date.humanize(locale=self.locale_fmt),
                )
                e["date"] = evt_date
                # translate event kind
                e["kind"] = self.isogeo_tr("events", e.get("kind"))

        # ---- IDENTIFICATION # ----------------------------------------------
        # Resource type
        resource_type = self.isogeo_tr("formatTypes", md.type)

        # Format
        if md.format and md.type in ("rasterDataset", "vectorDataset"):
            format_version = "{0} {1} ({2})".format(
                format_lbl,
                md.formatVersion,
                md.encoding,
            )

        # path to the resource
        localplace = md.path

        # ---- HISTORY # -----------------------------------------------------
        # data events
        if md.created:
            data_created = utils.hlpr_datetimes(md.created).strftime(self.dates_fmt)
        else:
            data_created = ""
        if md.modified:
            data_updated = utils.hlpr_datetimes(md.modified).strftime(self.dates_fmt)
        else:
            data_updated = ""
        if md.published:
            data_published = utils.hlpr_datetimes(md.published).strftime(self.dates_fmt)
        else:
            data_published = ""

        # validity
        if md.validFrom:
            validFrom = utils.hlpr_datetimes(md.validFrom).strftime(self.dates_fmt)
        else:
            validFrom = "NR"
        # end validity date
        if md.validTo:
            validTo = utils.hlpr_datetimes(md.validTo).strftime(self.dates_fmt)
        else:
            validTo = "NR"

        # ---- SPECIFICATIONS # -----------------------------------------------
        if md.specifications:
            specs_out = self.fmt.specifications(md.specifications)

        # ---- CGUs # --------------------------------------------------------
        if md.conditions:
            cgus_out = self.fmt.conditions(md.conditions)

        # ---- LIMITATIONS # -------------------------------------------------
        if md.limitations:
            lims_out = self.fmt.limitations(md.limitations)

        # ---- METADATA # ----------------------------------------------------
        md_created = utils.hlpr_datetimes(md._created).strftime(self.dates_fmt)
        md_updated = utils.hlpr_datetimes(md._modified).strftime(self.dates_fmt)

        # FILLFULLING THE TEMPLATE #
        context = {
            # "varThumbnail": InlineImage(docx_template, md.thumbnail),
            "varTitle": self.clean_xml(md.title),
            "varAbstract": self.clean_xml(md.abstract),
            "varNameTech": md.name,
            "varCollectContext": self.clean_xml(md.collectionContext),
            "varCollectMethod": self.clean_xml(md.collectionMethod),
            "varDataDtCrea": data_created,
            "varDataDtUpda": data_updated,
            "varDataDtPubl": data_published,
            "varValidityStart": validFrom,
            "varValidityEnd": validTo,
            "validityComment": self.clean_xml(md.validityComment),
            "varFormat": format_version,
            "varGeometry": md.geometry,
            "varObjectsCount": md.features,
            "varKeywords": " ; ".join(li_motscles),
            "varKeywordsCount": len(li_motscles),
            "varType": resource_type,
            "varOwner": owner,
            "varScale": md.scale,
            "varTopologyInfo": self.clean_xml(md.topologicalConsistency),
            "varInspireTheme": " ; ".join(li_theminspire),
            "varInspireConformity": inspire_valid,
            "varLimitations": lims_out,
            "varCGUS": self.fmt.conditions(cgus_out),
            "varSpecifications": specs_out,
            "varContactsCount": len(md.contacts),
            "varContactsDetails": contacts_out,
            "varSRS": srs,
            "varPath": localplace,
            "varFieldsCount": len(fields),
            "varFields": fields_out,
            "varEventsCount": len(md.events),
            "varEvents": events,
            "varMdDtCrea": md_created,
            "varMdDtUpda": md_updated,
            "varMdDtExp": datetime.now().strftime("%a %d %B %Y (%Hh%M)"),
            "varViewOC": link_visu,
            "varEditAPP": link_edit,
        }

        # fillfull file
        try:
            docx_template.render(context)
            logger.info(
                "Vector metadata stored: {} ({})".format(md.title_or_name(slugged=1), md.get("_id"))
            )
        except etree.XMLSyntaxError as e:
            logger.error(
                "Invalid character in XML: {}. "
                "Any special character (<, <, &...)? Check: {}".format(e, link_edit)
            )
        except (UnicodeEncodeError, UnicodeDecodeError) as e:
            logger.error(
                "Encoding error: {}. "
                "Any special character (<, <, &...)? Check: {}".format(e, link_edit)
            )
        except Exception as e:
            logger.error("Unexpected error: {}. Check: {}".format(e, link_edit))

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

        """Clean string from special characters.

        source: http://stackoverflow.com/a/5843560
        """
        return unicode(substitute).join(char for char in input_str if char.isalnum())

    def clean_xml(self, invalid_xml: str, mode="soft", substitute="_"):
        """Clean string of XML invalid characters.

        source: http://stackoverflow.com/a/13322581/2556577
        """
        if invalid_xml is None:
            return ""
        # assumptions:
        #   doc = *( start_tag / end_tag / text )
        #   start_tag = '<' name *attr [ '/' ] '>'
        #   end_tag = '<' '/' name '>'
        ws = r"[ \t\r\n]*"  # allow ws between any token
        name = "[a-zA-Z]+"  # note: expand if necessary but the stricter the better
        attr = '{name} {ws} = {ws} "[^"]*"'  # note: fragile against missing '"'; no "'"
        start_tag = "< {ws} {name} {ws} (?:{attr} {ws})* /? {ws} >"
        end_tag = "{ws}".join(["<", "/", "{name}", ">"])
        tag = "{start_tag} | {end_tag}"

        assert "{{" not in tag
        while "{" in tag:  # unwrap definitions
            tag = tag.format(**vars())

        tag_regex = re.compile("(%s)" % tag, flags=re.VERBOSE)

        # escape &, <, > in the text
        iters = [iter(tag_regex.split(invalid_xml))] * 2
        pairs = zip_longest(*iters, fillvalue="")  # iterate 2 items at a time

        # get the clean version
        clean_version = "".join(escape(text) + tag for text, tag in pairs)
        if mode == "strict":
            clean_version = re.sub(r"<.*?>", substitute, clean_version)
        else:
            pass
        return clean_version


# ###############################################################################
# ###### Stand alone program ########
# ###################################
if __name__ == "__main__":
    """
        Standalone execution and tests
    """
    # ------------ Specific imports ---------------------
    from ConfigParser import SafeConfigParser  # to manage options.ini
    from os import path

    # ------------ Settings from ini file ----------------
    if not path.isfile(path.realpath(r"..\settings_dev.ini")):
        logger.error(
            "To execute this script as standalone,"
            " you need to store your Isogeo application settings"
            " in a isogeo_params.ini file. You can use the template"
            " to set your own."
        )
        raise ValueError("settings.ini file missing.")
    else:
        pass

    config = SafeConfigParser()
    config.read(r"..\settings_dev.ini")

    settings = {s: dict(config.items(s)) for s in config.sections()}
    app_id = settings.get("auth").get("app_id")
    app_secret = settings.get("auth").get("app_secret")
    client_lang = settings.get("basics").get("def_codelang")

    # ------------ Connecting to Isogeo API ----------------
    # instanciating the class
    isogeo = Isogeo(client_id=app_id, client_secret=app_secret, lang="fr")

    token = isogeo.connect()

    # ------------ Isogeo search --------------------------
    search_results = isogeo.search(token, sub_resources=isogeo.sub_resources_available)

    # ------------ REAL START ----------------------------
    url_oc = "https://open.isogeo.com/s/c502e8f7c9da4c3aacdf3d905672d54c/Q4SvPfiIIslbdwkbWRFJLk7XWo4G0/"
    toDocx = Isogeo2docx()

    for md in search_results.get("results"):
        tpl = DocxTemplate(path.realpath(r"..\templates\template_Isogeo.docx"))
        toDocx.md2docx(tpl, md, url_oc)
        dstamp = datetime.now()
        if not md.get("name"):
            md_name = "NR"
        elif "." in md.get("name"):
            md_name = md.get("name").split(".")[1]
        else:
            md_name = md.get("name")
        tpl.save(
            r"..\output\{0}_{8}_{7}_{1}{2}{3}{4}{5}{6}.docx".format(
                "TestDemoDev",
                dstamp.year,
                dstamp.month,
                dstamp.day,
                dstamp.hour,
                dstamp.minute,
                dstamp.second,
                md.get("_id")[:5],
                md_name,
            )
        )
        del tpl
