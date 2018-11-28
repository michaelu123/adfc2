# encoding: utf-8

import os
import json
import logging
import time
import xml.sax
from myLogger import logger

weekdays = [ "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
character = [ "", "durchgehend Asphalt", "fester Belag", "unebener Untergrund", "unbefestigte Wege"]

def convertToMEZOrMSZ(beginning): # '2018-04-29T06:30:00+00:00'
    # scribus/Python2 does not support %z
    beginning = beginning[0:19] # '2018-04-29T06:30:00'
    d = time.strptime(beginning, "%Y-%m-%dT%H:%M:%S")
    oldDay = d.tm_yday
    if beginning.startswith("2017"):
        begSZ = "2017-03-26"
        endSZ = "2017-10-29"
        sz = beginning >= begSZ and beginning < endSZ
    elif beginning.startswith("2018"):
        begSZ = "2018-03-25"
        endSZ = "2018-10-28"
        sz = beginning >= begSZ and beginning < endSZ
    elif beginning.startswith("2019"):
        begSZ = "2019-03-31"
        endSZ = "2019-10-27"
        sz = beginning >= begSZ and beginning < endSZ
        """ Zeitumstellung wird eh 2020 abgeschafft!?
        elif beginning.startswith("2020"):
            begSZ = "2020-03-29"
            endSZ = "2020-10-25"
            sz = beginning >= begSZ and beginning < endSZ
        """
    else:
        raise ValueError("year " + beginning + " not configured")
    epochGmt = time.mktime(d)
    epochMez = epochGmt + ((2 if sz else 1) * 3600)
    mezTuple = time.localtime(epochMez)
    newDay = mezTuple.tm_yday
    mez = time.strftime("%Y-%m-%dT%H:%M:%S", mezTuple)
    if oldDay != newDay:
        logger.warning("day rollover from %s to %s", beginning, mez)
    return mez

class SAXHandler(xml.sax.handler.ContentHandler):
    def __init__(self):
        self.r = []
    def startElement(self, name, attrs):
        pass
    def endElement(self, name):
        pass
    def characters(self, content):
        self.r.append(content)
    def ignorableWhiteSpace(self, whitespace):
        pass
    def skippedEntity(self, name):
        pass
    def val(self):
        return "".join(self.r)

def removeHTML(s):
    if s.find("</") == -1:  # no HTML
        return s
    try:
        htmlHandler = SAXHandler()
        xml.sax.parseString("<xxxx>" + s + "</xxxx>", htmlHandler)
        return htmlHandler.val()
    except:
        logger.exception("can not parse '%s'", s )
        return s

# Clean text
def normalizeText(t):
    '''Rip off blank paragraphs, double spaces, html tags, quotes etc.'''
    changed = True
    while changed:
        changed = False
        t = t.strip()
        while t.count('***'):
            t = t.replace('***', '**')
            changed = True
        while t.count('**'):
            t = t.replace('**', '')
            changed = True
        while t.count('###'):
            t = t.replace('###', '##')
            changed = True
        while t.count('##'):
            t = t.replace('##', '')
            changed = True
        while t.count('~~~'):
            t = t.replace('~~~', '~~')
            changed = True
        while t.count('~~'):
            t = t.replace('~~', '')
            changed = True
        while t.count('\t'):
            t = t.replace('\t', ' ')
            changed = True
        if isinstance(t, str):  # crashes with Unicode/Scribus ??
            while t.count('\xa0'):
                t = t.replace('\xa0', ' ')
                changed = True
        while t.count('  '):
            t = t.replace('  ', ' ')
            changed = True
        while t.count('<br>'):
            t = t.replace('<br>', '\n')
            changed = True
        while t.count('\r'):  # DOS/Windows paragraph end.
            t = t.replace('\r', '\n')  # Change by new line
            changed = True
        while t.count('\n> '):
            t = t.replace('\n> ', '\n')
            changed = True
        while t.count(' \n'):
            t = t.replace(' \n', '\n')
            changed = True
        while t.count('\n '):
            t = t.replace('\n ', '\n')
            changed = True
        while t.count('\n\n'):
            t = t.replace('\n\n', '\n')
            changed = True
        if t.startswith('> '):
            t = t.replace('> ', '')
            changed = True
    return t

class Tour:
    def __init__(self, tourJS, tourJSSearch, tourServer):
        self.tourJS = tourJS
        self.tourJSSearch = tourJSSearch
        self.tourServer = tourServer
        self.tourLocations = tourJS.get("tourLocations")
        self.itemTags = tourJS.get("itemTags")
        self.eventItem = tourJS.get("eventItem")
        self.titel = self.eventItem.get("title").strip()
        logger.info("eventItemId %s", self.eventItem.get("eventItemId"))

    def getTitel(self):
        return self.titel
    def getFrontendLink(self):
        return "https://touren-termine.adfc.de/radveranstaltung/" + self.eventItem.get("cSlug")
    def getBackendLink(self):
        return "https://intern-touren-termine.adfc.de/modules/events/" + self.eventItem.get("eventItemId")

    def getNummer(self):
        num = self.tourJSSearch.get("tourNummer")
        if num is None:
            num = "999"
        return num

    def getAbfahrten(self):
        abfahrten = []
        for tourLoc in self.tourLocations:
            type = tourLoc.get("type")
            logger.debug("type %s", type)
            if type != "Startpunkt" and type != "Treffpunkt":
                continue
            if not tourLoc.get("withoutTime"):
                if len(abfahrten) == 0:  # for first loc, get starttime from eventItem, beginning in tourloc is often wrong
                    beginning = self.getDatum()[1]
                else:
                    beginning = tourLoc.get("beginning")
                    logger.debug("beginning %s", beginning) # '2018-04-24T12:00:00'
                    beginning = convertToMEZOrMSZ(beginning) # '2018-04-24T14:00:00'
                    beginning = beginning[11:16] # 14:00
            else:
                beginning = ""
            name = tourLoc.get("name")
            street = tourLoc.get("street")
            city = tourLoc.get("city")
            logger.debug("name '%s' street '%s' city '%s'", name, street, city)
            loc = name
            if city != "":
                if loc == "":
                    loc = city
                else:
                    loc = loc + " " + city
            if street != "":
                if loc == "":
                    loc = street
                else:
                    loc = loc + " " + street
            if type == "Startpunkt":
                if self.isTermin():
                    type = "Treffpunkt"
                else:
                    type = "Start"
            abfahrt = (type, beginning, loc)
            abfahrten.append(abfahrt)
        return abfahrten

    def getBeschreibung(self, removeNL):
        desc = self.eventItem.get("description")
        desc = normalizeText(desc)
        desc = removeHTML(desc)
        return desc

    def getKurzbeschreibung(self):
        desc = self.eventItem.get("cShortDescription")
        desc = normalizeText(desc)
        return desc

    def isTermin(self):
        return self.eventItem.get("eventType") == "Termin"

    def getSchwierigkeit(self):
        if self.isTermin():
            return "-"
        schwierigkeit = self.eventItem.get("cTourDifficulty")
        # apparently either 0 or between 1.0 and 5.0
        i = int(schwierigkeit + 0.5)
        return i # ["unbekannt", "sehr einfach, "einfach", "mittel", "schwer", "sehr schwer"][i] ??

    """ 
    itemtags has categories
        for Termine:
        "Aktionen, bei denen Rad gefahren wird" : getKategorie, e.g. Fahrrad-Demo, Critical Mass
        "Radlertreff / Stammtisch / \u00f6ffentliche Arbeits..." : getKategorie, e.g. Stammtisch
        "Serviceangebote": getKategorie, e.g. Codierung, Selbsthilfewerkstatt
        "Versammlungen" : getKategorie, e.g. Aktiventreff, Mitgliederversammlung
        "Vortr\u00e4ge & Kurse": getKategorie, e.g. Kurse, Radreisevortrag
        for Touren:
        "Besondere Charakteristik /Thema": getZusatzInfo
        "Besondere Zielgruppe" : getZusatzInfo
        "Geeignet f\u00fcr": getRadTyp
        "Typen (nach Dauer und Tageslage)" : getKategorie, e.g. Ganztagstour
        "Weitere Eigenschaften"  : getZusatzinfo, e.g. Bahnfahrt
    """

    def getMerkmale(self):
        merkmale = []
        for itemTag in self.itemTags:
            tag = itemTag.get("tag")
            merkmale.append(tag)
        return merkmale

    def getKategorie(self):
        for itemTag in self.itemTags:
            tag = itemTag.get("tag")
            category = itemTag.get("category")
            if category.startswith("Aktionen,") or category.startswith("Radlertreff") or category.startswith("Service") \
                    or category.startswith("Versammlungen") or category.startswith("Vortr")\
                    or category.startswith("Typen ")                     :
                return tag.split()[0]
        return "Ohne"

    def getRadTyp(self):
        # wenn nur Rennrad oder nur Mountainbike, dann dieses, sonst Tourenrad
        l = 0
        for itemTag in self.itemTags:
            category = itemTag.get("category")
            if category.startswith("Geeignet "):
                l += 1
        for itemTag in self.itemTags:
            tag = itemTag.get("tag")
            category = itemTag.get("category")
            if category.startswith("Geeignet "):
                if l == 1 and (tag == "Rennrad" or tag == "Mountainbike"):
                    return tag
        return "Tourenrad"

    def getZusatzInfo(self):
        besonders = []
        weitere = []
        zielgruppe = []
        for itemTag in self.itemTags:
            tag = itemTag.get("tag")
            category = itemTag.get("category")
            if category == "Besondere Charakteristik /Thema":
                besonders.append(tag)
            if category == "Weitere Eigenschaften":
                weitere.append(tag)
            if category == "Besondere Zielgruppe":
                zielgruppe.append(tag)
        zusatzinfo = []
        if len(besonders) > 0:
            besonders = "Besondere Charakteristik/Thema: " + ", ".join(besonders)
            zusatzinfo.append(besonders)
        if len(weitere) > 0:
            weitere = "Weitere Eigenschaften: " + ", ".join(weitere)
            zusatzinfo.append(weitere)
        if len(zielgruppe) > 0:
            zielgruppe = "Besondere Zielgruppe: " + ", ".join(zielgruppe)
            zusatzinfo.append(zielgruppe)
        return zusatzinfo


    def getStrecke(self):
        l = self.eventItem.get("cTourLengthKm")
        return str(l) + " km"

    def getHoehenmeter(self):
        h = self.eventItem.get("cTourHeight")
        return str(h)

    def getCharacter(self):
        c = self.eventItem.get("cTourSurface")
        return character[c]

    def getDatum(self):
        datum = self.eventItem.get("beginning")
        datum = convertToMEZOrMSZ(datum)
        # fromisoformat defined in Python3.7, not used by Scribus
        # date = datetime.fromisoformat(datum)
        logger.debug("datum <%s>", str(datum))
        day = str(datum[0:10])
        date = time.strptime(day, "%Y-%m-%d")
        weekday = weekdays[date.tm_wday]
        res =  (weekday + ", " + day[8:10] + "." + day[5:7] + "." + day[0:4], datum[11:16])
        return res

    def getDatumRaw(self):
        return self.eventItem.get("beginning")

    def getEndDatum(self):
        enddatum = self.eventItem.get("end")
        enddatum = convertToMEZOrMSZ(enddatum)
        # fromisoformat defined in Python3.7, not used by Scribus
        # enddatum = datetime.fromisoformat(enddatum)
        logger.debug("enddatum %s", str(enddatum))
        day = str(enddatum[0:10])
        date = time.strptime(day, "%Y-%m-%d")
        weekday = weekdays[date.tm_wday]
        res = (weekday + ", " + day[8:10] + "." + day[5:7] + "." + day[0:4], enddatum[11:16])
        return res

    def getEndDatumRaw(self):
        return self.eventItem.get("end")

    def getPersonen(self):
        personen = []
        organizer = self.eventItem.get("cOrganizingUserId")
        if organizer is not None and len(organizer) > 0:
            org = self.tourServer.getUser(organizer)
            if org != None:
                personen.append(str(org))
        organizer2 = self.eventItem.get("cSecondOrganizingUserId")
        if organizer2 is not None and len(organizer2) > 0 and organizer2 != organizer:
            org = self.tourServer.getUser(organizer2)
            if org != None:
                personen.append(str(org))
        return personen

    def getImagePreview(self):
        return self.tourJS.get("imagePreview")

    def getName(self):
        tourLoc = self.tourLocations[0]
        return tourLoc.get("name")

    def getCity(self):
        tourLoc = self.tourLocations[0]
        return tourLoc.get("city")

    def getStreet(self):
        tourLoc = self.tourLocations[0]
        return tourLoc.get("street")


class User:
    def __init__(self, userJS):
        u = userJS.get("user")
        self.firstName = u.get("firstName")
        self.lastName = u.get("lastName")
        try:
            self.phone = u.get("cellPhone")
            if self.phone == None or self.phone == "":
                self.phone = userJS.get("temporaryContacts")[0].get("phone")
        except:
            self.phone = None
    def __repr__(self):
        name = self.firstName + " " + self.lastName
        if self.phone != None and self.phone != "":
            name += " (" + self.phone + ")"
        return name
