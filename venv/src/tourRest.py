# encoding: utf-8

import os
import json
import logging
import time
import xml.sax
from myLogger import logger

weekdays = [ "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]

def convertToMEZOrMSZ(beginning): # '2018-04-29T06:30:00+00:00'
    # scribus/Python2 does not support %z
    beginning = beginning[0:19] # '2018-04-29T06:30:00'
    d = time.strptime(beginning, "%Y-%m-%dT%H:%M:%S")
    oldDay = d.tm_yday
    if beginning.startswith("2018"):
        begSZ = "2018-03-25"
        endSZ = "2018-10-28"
        sz = beginning >= begSZ and beginning < endSZ
    elif beginning.startswith("2019"):
        begSZ = "2019-03-31"
        endSZ = "2019-10-27"
        sz = beginning >= begSZ and beginning < endSZ
    elif beginning.startswith("2020"):
        begSZ = "2019-03-29"
        endSZ = "2019-10-25"
        sz = beginning >= begSZ and beginning < endSZ
    else:
        raise ValueError("year " + beginning + " not configured")
    epochGmt = time.mktime(d)
    epochMez = epochGmt + ((2 if sz else 1) * 3600)
    mezTuple = time.localtime(epochMez)
    newDay = mezTuple.tm_yday
    mez = time.strftime("%Y-%m-%dT%H:%M:%S", mezTuple)
    if oldDay != newDay:
        raise ValueError("day rollover for tour %s from %s to %s", titel, beginning, mez)
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
    if s.find("<span") == -1:  # no HTML
        return s
    htmlHandler = SAXHandler()
    xml.sax.parseString("<xxxx>" + s + "</xxxx>", htmlHandler)
    return htmlHandler.val()

class Tour:
    def __init__(self, tourJS):
        self.tourJS = tourJS
        self.tourLocations = tourJS.get("tourLocations")
        self.itemTags = tourJS.get("itemTags")
        self.eventItem = tourJS.get("eventItem")

    def getTitel(self):
        return self.eventItem.get("title")

    def getAbfahrten(self):
        abfahrten = []
        for tourLoc in self.tourLocations:
            type = tourLoc.get("type")
            logger.debug("type %s", type)
            if type != "Startpunkt" and type != "Treffpunkt":
                continue
            beginning = tourLoc.get("beginning")
            logger.debug("beginning %s", beginning) # '2018-04-24T12:00:00'
            beginning = convertToMEZOrMSZ(beginning) # '2018-04-24T14:00:00'
            beginning = beginning[11:16] # 14:00
            name = tourLoc.get("name")
            street = tourLoc.get("street")
            city = tourLoc.get("city")
            logger.debug("name %s street %s city %s", name, street, city)
            loc = name + " " + city + ", " + street
            abfahrt = (beginning, loc)
            abfahrten.append(abfahrt)
        return abfahrten

    def getBeschreibung(self):
        desc = self.eventItem.get("description").replace("\n", " ")
        desc = removeHTML(desc)
        return desc

    def getSchwierigkeit(self):
        kategorie = self.getKategorie()
        if kategorie == "Feierabendtour":
            return "F"
        if kategorie == "Halbtagestour" or kategorie == "Tagestour" or kategorie == "Mehrtagestour":
            schwierigkeit = self.eventItem.get("cTourDifficulty")
            return "*" * int(schwierigkeit + 0.5)  # ???
        else:
            raise ValueError("Unbekannte Kategorie " + kategorie)

    def getKategorie(self):
        for itemTag in self.itemTags:
            tag = itemTag.get("tag")
            category = itemTag.get("category")
            if category == "Typen (nach Dauer und Tageslage)":
                return tag
        raise ValueError("Keine Kategorie definiert (z.B. Feierabendtour, Halbtagstour...)")

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
        if len(besonders) > 0:
            besonders = "Besondere Charakteristik/Thema: " + ", ".join(besonders)
        else:
            besonders = ""
        if len(weitere) > 0:
            weitere = "Weitere Eigenschaften: " + ", ".join(weitere)
        else:
            weitere = ""
        if len(zielgruppe) > 0:
            zielgruppe = "Besondere Zielgruppe: " + ", ".join(zielgruppe)
        else:
            zielgruppe = ""
        return [besonders, weitere, zielgruppe]

    def getStrecke(self):
        return str(self.eventItem.get("cTourLengthKm")) + " km"

    def getDatum(self):
        datum = self.eventItem.get("beginning")
        beginning = convertToMEZOrMSZ(datum)
        # fromisoformat defined in Python3.7, not used by Scribus
        # date = datetime.fromisoformat(datum)
        datum = str(datum[0:10])
        logger.debug("datum <%s> ", str(datum))
        date = time.strptime(datum, "%Y-%m-%d")
        weekday = weekdays[date.tm_wday]
        res =  weekday + ", " + datum[8:10] + "." + datum[5:7] + "." + datum[0:4]
        return res

    def getEndDatum(self):
        enddatum = self.eventItem.get("end")
        enddatum = convertToMEZOrMSZ(enddatum)
        # fromisoformat defined in Python3.7, not used by Scribus
        # enddate = datetime.fromisoformat(enddatum)
        enddatum = str(enddatum[0:10])
        #enddate = datetime.strptime(enddatum, "%Y-%m-%d")
        enddate = time.strptime(enddatum, "%Y-%m-%d")
        logger.debug("enddate %s %s ", type(enddate), str(enddate))
        weekday = weekdays[enddate.tm_wday]
        return weekday + ", " + enddatum[8:10] + "." + enddatum[5:7] + "." + enddatum[0:4]

    def getPersonen(self):
        """
              Starnberg:
                    <Leiter>
                        <Person>
                          <Name>Martin Held</Name>
                        </Person>
                        <Person>
                          <Name>Claus Piesch</Name>
                        </Person>
                        <Person>
                          <Name>Jochen Twiehaus</Name>
                        </Person>
                        <Person>
                          <Name><TelMobil>0171/2755036</TelMobil>Anton Maier</Name>
                        </Person>
                    </Leiter>

              tour.getElementsByTagName("Leiter")[0].getElementsByTagName("Person")
              Json:

        """
        personen = []
        organizer = self.eventItem.get("cOrganizingUserId")
        if organizer is not None and len(organizer) > 0:
            personen.append(organizer)
        organizer2 = self.eventItem.get("cSecondOrganizingUserId")
        if organizer2 is not None and len(organizer2) > 0:
            if organizer2 != organizer:
                personen.append(organizer2)
        if len(personen) == 0:
            logger.error("Tour %s hat keinen Tourleiter", titel )
        return personen
