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
import os
import tourRest

logging.basicConfig(level=logging.DEBUG, filename = "adfc_rest1.log", filemode="w")
logger = logging.getLogger("adfc-rest1")
logger.info("cwd=%s", os.getcwd())

keySta = "152085"
keyGau = "15208514"
keyMuc = "152059"

#useRest = False
#unitKey = keySta
#tpConn = None

logger.info("1sta %s", __name__)
try:
    import scribus
    import httplib  # scribus seems to use Python 2
    useScribus = True
    unitKey = scribus.valueDialog("Gliederung", "Bitte Nummer der Gliederung angeben")
    yOrN = scribus.valueDialog("UseRest", "Sollen aktuelle Daten vom Server geholt werden? (j/n)").lower()[0]
    useRest = yOrN == 'j' or yOrN == 'y' or yOrN == 't'
    from forPy2 import myprint
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
    from forPy3 import myprint


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

def handleTour(eventItemId):
    global tpConn
    jsonPath = "c:/temp/tpjson/" + eventItemId + ".json"
    if not os.path.exists(jsonPath) or useRest:
        if tpConn is None:
            tpConn = httplib.HTTPSConnection("api-touren-termine.adfc.de")
        tpConn.request("GET", "/api/eventItems/" + eventItemId)
        resp = tpConn.getresponse()
        logger.debug("resp %d %s", resp.status, resp.reason)
        tourJS = json.load(resp)
        tourJS["eventItemFiles"] = None #save space
        # if not os.path.exists(jsonPath):
        with open(jsonPath, "w") as jsonFile:
            json.dump(tourJS, jsonFile, indent=4)
    else:
        with open(jsonPath, "r") as jsonFile:
            tourJS = json.load(jsonFile)

    tour = tourRest.Tour(logger, tourJS)
    try:
        titel = tour.getTitel()
        logger.info("Title %s", titel)
        datum = tour.getDatum()
        logger.info("datum %s", datum)

        abfahrten = tour.getAbfahrten()
        if len(abfahrten) == 0:
            raise ValueError("kein Startpunkt in tour %s", titel)
            return
        logger.info("abfahrten %s ", str(abfahrten))

        beschreibung = tour.getBeschreibung()
        logger.info("beschreibung %s", beschreibung)
        zusatzinfo = tour.getZusatzInfo()
        logger.info("zusatzinfo %s", str(zusatzinfo))
        kategorie = tour.getKategorie()
        logger.info("kategorie %s", kategorie)
        schwierigkeit = tour.getSchwierigkeit()
        logger.info("schwierigkeit %s", schwierigkeit)
        strecke = tour.getStrecke()
        logger.info("strecke %s", strecke)

        if kategorie == 'Mehrtagestour':
            enddatum = tour.getEndDatum()
            logger.info("enddatum %s", enddatum)

        personen = tour.getPersonen()
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




