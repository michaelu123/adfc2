#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
"""
USAGE
You must have a document open, and a text frame selected.
Simply run the script, which asks if you want to read from the rest interface or from
(previously written) json files.
Note that any text in the frame will be deleted before the text from
the XML file is added.
The script also assumes you have created a styles named 'Radtour_titel' etc, which
it will apply to the frame.
"""
import json
import logging
import sys
import time
import locale
import os
import xml.sax

logging.basicConfig(level=logging.DEBUG, filename = "adfc-rest1.log", filemode="w")
logger = logging.getLogger("adfc-rest1")
logger.info("cwd=%s", os.getcwd())

# URL der Touren des KV Starnberg https://api-touren-termine.adfc.de/api/eventItems/search?unitKey=152085
urlSta = "https://api-touren-termine.adfc.de/api/eventItems/search?unitKey="
keySta = "152085"
keyGau = "15208514"
keyMuc = "152059"
weekdays = [ "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]

useRest = False
unitKey = keySta
tpConn = None

logger.info("1sta %s", __name__)
try:
    import scribus
    import httplib  # scribus seems to use Python 2
    useScribus = True
    unitKey = scribus.valueDialog("Gliederung", "Bitte Nummer der Gliederung angeben")
    yOrN = scribus.valueDialog("UseRest", "Sollen aktuelle Daten vom Server geholt werden? (j/n)").lower()[0]
    useRest = yOrN == 'j' or yOrN == 'y' or yOrN == 't'
except ImportError:
    import http.client as httplib
    import argparse
    useScribus = False
    parser = argparse.ArgumentParser(description="Formatiere Daten des Tourenportals")
    parser.add_argument("nummer", help="Gliederungsnummer, z.B. 152059 für München")
    parser.add_argument("useRest", help="Sollen aktuelle Daten vom Server geholt werden? (j/n)")
    args = parser.parse_args()
    unitKey = args.nummer
    yOrN = args.useRest.lower()[0]
    useRest = yOrN == 'j' or yOrN == 'y' or yOrN == 't'

def myprint(*a):
    pass #print (*a, end='')

class ScribusTest: # just log scribus calls
    def setStyle(self, style, frame):
        logger.info("scribus setStyle style %s frame %s", str(style), str(frame))
    def createParagraphStyle(self, style):
        logger.info("scribus createParagraphStyle %s", str(style))
    def selectText(self, len, n, textbox):
        logger.info("scribus selectText %d %d %s", len, n, str(textbox))
    def deleteText(self, textbox):
        logger.info("scribus deleteText %s", str(textbox))
    def insertText(self, text, off, textbox):
        logger.info("scribus insertText %s %d %s", text, off, str(textbox))
    def messageBox(self, text1, text2, lvl, button):
        logger.info("scribus.messagetext %s %s %d %s", text1, text2, lvl, str(button))
    def haveDoc(self):
        return True
    def selectionCount(self):
        return 1
    def getSelectedObject(self):
        return None
    def getObjectType(self, textbox):
        return "TextFrame"
    def getTextLength(self):
        return 0

class DatenTest: # just log data
    def setStyle(self, style, frame):
        pass
    def createParagraphStyle(self, style):
        pass
    def selectText(self, len, n, textbox):
        pass
    def deleteText(self, textbox):
        pass
    def insertText(self, text, off, textbox):
        logger.info(text)
        myprint (text)
    def messageBox(self, text1, text2, lvl, button):
        pass
    def haveDoc(self):
        return True
    def selectionCount(self):
        return 1
    def getSelectedObject(self):
        return None
    def getObjectType(self, textbox):
        return "TextFrame"
    def getTextLength(self):
        return 0

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

logger.info("2sta %s", __name__)
if not useScribus:
    #scribus = ScribusTest()
    scribus = DatenTest()

def addStyle(style, frame):
     try:
         scribus.setStyle(style, frame)
     except scribus.NotFoundError:
         scribus.createParagraphStyle(style)
         scribus.setStyle(style, frame)

def handleAbfahrt(abfahrt):
    # abfahrt = (beginning, loc)
    uhrzeit = abfahrt[0]
    ort = abfahrt[1]
    logger.info("Abfahrt: uhrzeit=%s ort=%s", uhrzeit, ort)
    scribus.setStyle('Radtour_start',textbox)
    scribus.insertText('Start: '+uhrzeit+', '+ort+'\n',-1,textbox)


def handleTextfeld(stil,textelement):
    logger.info("Textfeld: stil=%s text=%s", stil, textelement)
    if textelement != None:
        scribus.setStyle(stil,textbox)
        scribus.insertText(textelement+'\n',-1,textbox)

def handleTextfeldList(stil,textList):
    logger.info("TextfeldList: stil=%s text=%s", stil, str(textList))
    for text in textList:
        if len(text) == 0:
            continue
        logger.info("Text: stil=%s text=%s", stil, text)
        scribus.setStyle(stil,textbox)
        scribus.insertText(text+'\n',-1,textbox)

def handleBeschreibung(textelement):
    handleTextfeld(textelement)

def handleTel(Name):
    telfestnetz = Name.getElementsByTagName("TelFestnetz")
    telmobil = Name.getElementsByTagName("TelMobil")
    if len(telfestnetz)!=0:
        logger.info("Tel: festnetz=%s", telfestnetz[0].firstChild.data)
        scribus.insertText(' ('+telfestnetz[0].firstChild.data+')',-1,textbox)
    if len(telmobil)!=0:
        logger.info("Tel: mobil=%s", telmobil[0].firstChild.data)
        scribus.insertText(' ('+telmobil[0].firstChild.data+')',-1,textbox)

def handleName(name):
    logger.info("Name: name=%s", name)
    scribus.insertText(name,-1,textbox)
    # handleTel(name) ham wer nich!

def handleTourenleiter(TLs):
    scribus.setStyle('Radtour_tourenleiter',textbox)
    scribus.insertText('Tourenleiter: ',-1,textbox)
    names = ", ".join(TLs)
    scribus.insertText(names,-1,textbox)
    scribus.insertText('\n',-1,textbox)

def handleTitel(tt):
    logger.info("Titel: titel=%s", tt)
    scribus.setStyle('Radtour_titel',textbox)
    scribus.insertText(tt+'\n',-1,textbox)

def handleKopfzeile(dat, kat, schwierig, strecke):
    logger.info("Kopfzeile: dat=%s kat=%s schwere=%s strecke=%s", dat, kat, schwierig, strecke)
    scribus.setStyle('Radtour_kopfzeile',textbox)
    scribus.insertText(dat+':	'+kat+'	'+schwierig+'	'+strecke+'\n',-1,textbox)

def handleKopfzeileMehrtage(anfang, ende, kat, schwierig, strecke):
    logger.info("Mehrtage: anfang=%s ende=%s kat=%s schwere=%s strecke=%s", anfang, ende, kat, schwierig, strecke)
    scribus.setStyle('Radtour_kopfzeile',textbox)
    scribus.insertText(anfang+' bis '+ende+':\n',-1,textbox)
    scribus.setStyle('Radtour_kopfzeile',textbox)
    scribus.insertText('	'+kat+'	'+schwierig+'	'+strecke+'\n',-1,textbox)

logger.info("3sta")

def getAbfahrten(tourLocations):
    abfahrten = []
    for tourLoc in tourLocations:
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

def removeHTML(s):
    if s.find("<span") == -1:  # no HTML
        return s
    htmlHandler = SAXHandler()
    xml.sax.parseString("<xxxx>" + s + "</xxxx>", htmlHandler)
    return htmlHandler.val()

def getBeschreibung(eventItem):
    desc = eventItem.get("description").replace("\n", " ")
    desc = removeHTML(desc)
    return desc

def getSchwierigkeit(eventItem, itemTags):
    kategorie = getKategorie(itemTags)
    if kategorie == "Feierabendtour":
        return "F"
    if kategorie == "Halbtagestour" or kategorie == "Tagestour" or kategorie == "Mehrtagestour":
        schwierigkeit = eventItem.get("cTourDifficulty")
        return "*" * int(schwierigkeit + 0.5)  # ???
    else:
        raise ValueError("Unbekannte Kategorie " + kategorie)

def getKategorie(itemTags):
    for itemTag in itemTags:
        tag = itemTag.get("tag")
        category = itemTag.get("category")
        if category == "Typen (nach Dauer und Tageslage)":
            return tag
    raise ValueError("Keine Kategorie definiert (z.B. Feierabendtour, Halbtagstour...)")

def getZusatzInfo(itemTags):
    besonders = []
    weitere = []
    zielgruppe = []
    for itemTag in itemTags:
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

def getStrecke(eventItem):
    return str(eventItem.get("cTourLengthKm")) + " km"

def getDatum(eventItem):
    datum = eventItem.get("beginning")
    beginning = convertToMEZOrMSZ(datum)
    # fromisoformat defined in Python3.7, not used by Scribus
    # date = datetime.fromisoformat(datum)
    datum = str(datum[0:10])
    logger.debug("datum <%s> ", str(datum))
    date = time.strptime(datum, "%Y-%m-%d")
    weekday = weekdays[date.tm_wday]
    res =  weekday + ", " + datum[8:10] + "." + datum[5:7] + "." + datum[0:4]
    return res

def getEndDatum(eventItem):
    enddatum = eventItem.get("end")
    enddatum = convertToMEZOrMSZ(enddatum)
    # fromisoformat defined in Python3.7, not used by Scribus
    # enddate = datetime.fromisoformat(enddatum)
    enddatum = str(enddatum[0:10])
    #enddate = datetime.strptime(enddatum, "%Y-%m-%d")
    enddate = time.strptime(enddatum, "%Y-%m-%d")
    logger.debug("enddate %s %s ", type(enddate), str(enddate))
    weekday = weekdays[enddate.tm_wday]
    return weekday + ", " + enddatum[8:10] + "." + enddatum[5:7] + "." + enddatum[0:4]

def getPersonen(eventItem):
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
    organizer = eventItem.get("cOrganizingUserId")
    if organizer is not None and len(organizer) > 0:
        personen.append(organizer)
    organizer2 = eventItem.get("cSecondOrganizingUserId")
    if organizer2 is not None and len(organizer2) > 0:
        if organizer2 != organizer:
            personen.append(organizer2)
    if len(personen) == 0:
        logger.error("Tour %s hat keinen Tourleiter", titel )
    return personen

def handleTour(eventItemId):
    global tpConn
    jsonPath = "c:/temp/tpjson/" + eventItemId + ".json"
    if not os.path.exists(jsonPath) or useRest:
        if tpConn is None:
            tpConn = httplib.HTTPSConnection("api-touren-termine.adfc.de")
        tpConn.request("GET", "/api/eventItems/" + eventItemId)
        resp = tpConn.getresponse()
        logger.debug("resp %d %s", resp.status, resp.reason)
        tour = json.load(resp)
        tour["eventItemFiles"] = None #save space
        # if not os.path.exists(jsonPath):
        with open(jsonPath, "w") as jsonFile:
            json.dump(tour, jsonFile, indent=4)
    else:
        with open(jsonPath, "r") as jsonFile:
            tour = json.load(jsonFile)

    tourLocations = tour.get("tourLocations")
    itemTags = tour.get("itemTags")
    eventItem = tour.get("eventItem")
    try:
        titel = eventItem.get("title")
        logger.info("Title %s", titel)
        datum = getDatum(eventItem)
        logger.info("datum %s", datum)

        abfahrten = getAbfahrten(tourLocations)
        if len(abfahrten) == 0:
            raise ValueError("kein Startpunkt in tour %s", titel)
            return
        logger.info("abfahrten %s ", str(abfahrten))

        beschreibung = getBeschreibung(eventItem)
        logger.info("beschreibung %s", beschreibung)
        zusatzinfo = getZusatzInfo(itemTags)
        logger.info("zusatzinfo %s", str(zusatzinfo))
        kategorie = getKategorie(itemTags)
        logger.info("kategorie %s", kategorie)
        schwierigkeit = getSchwierigkeit(eventItem, itemTags)
        logger.info("schwierigkeit %s", schwierigkeit)
        strecke = getStrecke(eventItem)
        logger.info("strecke %s", strecke)

        if kategorie == 'Mehrtagestour':
            enddatum = getEndDatum(eventItem)
            logger.info("enddatum %s", enddatum)

        personen = getPersonen(eventItem)
        logger.info("personen %s", str(personen))
    except Exception as e:
        logger.error("Fehler in der Tour %s: %s", titel, e)
        myprint("\nFehler in der Tour ", titel, ": ", e)
        return

    scribus.insertText('\n',-1,textbox)
    if kategorie == 'Mehrtagestour':
        handleKopfzeileMehrtage(datum, enddatum, kategorie, schwierigkeit, strecke)
    else:
        handleKopfzeile(datum, kategorie, schwierigkeit, strecke)
    handleTitel(titel)
    for abfahrt in abfahrten:
        handleAbfahrt(abfahrt)
    handleTextfeld('Radtour_beschreibung',beschreibung)
    handleTourenleiter(personen)
    handleTextfeldList('Radtour_zusatzinfo',zusatzinfo)


if not scribus.haveDoc():
     scribus.messageBox('Scribus - Script Error', "No document open", scribus.ICON_WARNING, scribus.BUTTON_OK)
     sys.exit(1)

if scribus.selectionCount() == 0:
     scribus.messageBox('Scribus - Script Error',
             "There is no object selected.\nPlease select a text frame and try again.",
             scribus.ICON_WARNING, scribus.BUTTON_OK)
     sys.exit(2)
if scribus.selectionCount() > 1:
     scribus.messageBox('Scribus - Script Error',
             "You have more than one object selected.\nPlease select one text frame and try again.",
             scribus.ICON_WARNING, scribus.BUTTON_OK)
     sys.exit(2)

textbox = scribus.getSelectedObject()
ftype = scribus.getObjectType(textbox)

if ftype != "TextFrame":
     scribus.messageBox('Scribus - Script Error', "This is not a textframe. Try again.", scribus.ICON_WARNING, scribus.BUTTON_OK)
     sys.exit(2)

scribus.deleteText(textbox)
scribus.setStyle('Radtouren_titel',textbox)
scribus.insertText('Radtouren\n',0,textbox)

jsonPath = "c:/temp/tpjson/search-" + unitKey + ".json"
if not os.path.exists(jsonPath) or useRest:
    tpConn = httplib.HTTPSConnection("api-touren-termine.adfc.de")
    logger.debug("tpConn %s %s ", type(tpConn), str(tpConn))
    tpConn.request("GET", "/api/eventItems/search?unitKey=" + unitKey)
    resp = tpConn.getresponse()
    logger.debug("resp %s %s ", type(resp), str(resp))
    jsRoot = json.load(resp)
else:
    resp = None
    with open(jsonPath, "r") as jsonFile:
        jsRoot = json.load(jsonFile)

items = jsRoot.get("items")
touren = []
for item in iter(items):
    item["imagePreview"] = "" # save space
    titel = item.get("title")
    if titel is None:
        logger.error("Kein Titel für die Tour %s", str(item) )
        continue;
    if item.get("eventType") != "Radtour":
        continue;
    if item.get("beginning") is None:
        logger.error("Kein Beginn für die Tour %s", str(item) )
        continue;
    # add other filter conditions here
    touren.append(item)

if not resp is None: # and not os.path.exists(jsonPath):
    with open(jsonPath, "w") as jsonFile:
        json.dump(jsRoot, jsonFile, indent=4)

def tourdate(t):
    return t.get("beginning")

touren.sort(key=tourdate) # sortieren nach Datum
for tour in touren:
    eventItemId = tour.get("eventItemId");
    handleTour(eventItemId)




