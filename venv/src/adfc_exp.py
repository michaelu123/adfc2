#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
USAGE
You must have a document open, and a text frame selected.
Simply run the script, which asks for a XML file with a file dialog.
Note that any text in the frame will be deleted before the text from
the XML file is added.
The script also assumes you have created a styles named 'Radtour_titel' etc, which
it will apply to the frame.
"""

import sys, time, locale, os
import xml.dom.minidom
import logging
import locale

logging.basicConfig(level=logging.INFO) #, filename="adfc_exp.log")
logger = logging.getLogger("adfc-exp")
logger.info("cwd=%s name=%s", os.getcwd(), __name__)
weekdays = [ "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]

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

    def fileDialog(self, a, b):
        #return "Veranstaltungs-Export_Sep1_muh.xml"
        return "Veranstaltungs-Export_Sep1_all.xml"
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

    def fileDialog(self, a, b):
        #return "Veranstaltungs-Export_Sep1_muh.xml"
        return "Veranstaltungs-Export_Sep1_all.xml"
    def getTextLength(self):
        return 0

logger.info("1sta %s", __name__)
try:
     import scribus
except ImportError:
     print("Unable to import the 'scribus' module. This script will only run within")
     print("the Python interpreter embedded in Scribus. Try Script->Execute Script.")
     # sys.exit(1)
     #scribus = ScribusTest()
     scribus = DatenTest()

logger.info("2sta")

def addStyle(style, frame):
     try:
         scribus.setStyle(style, frame)
     except scribus.NotFoundError:
         scribus.createParagraphStyle(style)
         scribus.setStyle(style, frame)

def getText(nodelist):
    rc = []
    for node in nodelist:
        if node.nodeType == node.TEXT_NODE:
            rc.append(node.data)
    return ''.join(rc)

def backSpaceNChar(n):
    # delete n chars at the current position (end of text)
    logger.info("back %d", n)
    scribus.selectText(scribus.getTextLength()-n, n, textbox)
    scribus.deleteText(textbox)

def handleAbfahrt(abfahrt):
    # abfahrt = (beginning, loc)
    uhrzeit = abfahrt[0]
    ort = abfahrt[1]
    logger.info("Abfahrt: uhrzeit=%s ort=%s", uhrzeit, ort)
    scribus.setStyle('Radtour_start',textbox)
    scribus.insertText('Start: '+uhrzeit+', '+ort+'\n',-1,textbox)


def handleTextfeld(stil,textelement):
    """ Starnberg
        absaetze = textelement.getElementsByTagName("p")
        for absatz in absaetze:
            logger.info("Textfeld: stil=%s text=%s", stil, absatz.childNodes[0].data)
            scribus.setStyle(stil,textbox)
            scribus.insertText(absatz.childNodes[0].data+'\n',-1,textbox)
        Neu: einfacher String
    """
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
    for person in TLs:
        handleName(person)
        scribus.insertText(', ',-1,textbox)
    # Lösche die letzten beiden Zeichen ', ' wieder und hänge den Zeilenumbruch an.
    backSpaceNChar(2)
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

def tourdate(tour):
    return tour.getElementsByTagName("Beginning")[0].firstChild.data

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

xmlfiledata = scribus.fileDialog('XML file', 'XML files(*.xml)')

DomTree = xml.dom.minidom.parse(xmlfiledata)
collection = DomTree.documentElement

# Alle Touren einlesen
touren = collection.getElementsByTagName("ExportEventItem")
touren.sort(key = tourdate)
#logger.info("touren %s %s ", type(touren), str(touren)) # too large

# die Touren ausgeben

def getAbfahrten():
    abfahrten = []
    startPunkt = ""

    tourLocs = tour.getElementsByTagName("TourLocations")
    logger.debug("tls %s %s ", type(tourLocs), str(tourLocs))
    if len(tourLocs) != 1:
        logger.error("size(tourlocs=%d", size(tourLocs))
    tourLoc = tourLocs[0]
    logger.debug("tl %s %s ", type(tourLoc), str(tourLoc))
    expTourLocs = tourLoc.getElementsByTagName("ExportTourLocation")
    logger.debug("etls %s %d %s ", type(expTourLocs), len(expTourLocs), str(expTourLocs))

    for expTourLoc in expTourLocs:
        logger.debug("etl %s %s ", type(expTourLocs), str(expTourLocs))
        type1 = expTourLoc.getElementsByTagName("Type")
        logger.debug("type1 %s %s ", type(type1), str(type1))
        type2 = type1[0].firstChild.data
        logger.debug("type2 %s %s ", type(type2), str(type2))
        if type2 != "Startpunkt" and type2 != "Treffpunkt":
            continue
        if type2 == "Startpunkt":
            startPunkt = type2
        beginning = expTourLoc.getElementsByTagName("Beginning")[0].firstChild.data
        logger.debug("beginning %s %s ", type(beginning), str(beginning)) # '2018-04-24T12:00:00'
        beginning = convertToMEZOrMSZ(beginning) # '2018-04-24T14:00:00'
        beginning = beginning[11:16] # 14:00
        loc = expTourLoc.getElementsByTagName("Name")[0].firstChild
        logger.debug("Name %s %s ", type(loc), str(loc))
        if loc is None:
            loc = expTourLoc.getElementsByTagName("Street")[0].firstChild
            logger.info("Street %s %s ", type(loc), str(loc))
        loc = loc.data
        logger.debug("loc %s %s ", type(loc), str(loc))
        abfahrt = (beginning, loc)
        abfahrten.append(abfahrt)
    return abfahrten

def convertToMEZOrMSZ(beginning):
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
    mez = time.localtime(epochMez)
    newDay = mez.tm_yday
    s = time.strftime("%Y-%m-%dT%H:%M:%S", mez)
    if oldDay != newDay:
        raise ValueError("day rollover for tour %s from %s to %s", titel, beginning, s)
    return s

def getBeschreibung():
    """  Beschreibung:
        Starnberg:
          <Beschreibung>
            <p>Tutzing-Tour zu neuralgischen Punkten mit Verbesserungspotenzial:</p>
            <p>Stationen u.a. Bahnhof, Geschäftszentrum Lindemannstraße, Einmündung Lindemannstraße – Hauptstraße, Schulen, Gestaltung Ortsmitte, Strecke bis Garatshausen, Wohnviertel.</p>
            <p>Gegen 12.00 Uhr enden wir am Tutzinger Hof zur Gründung der ADFC Ortsgruppe Tutzing.</p>
            <p>Anschließend Ausklang im Biergarten.</p>
          </Beschreibung>
        Neu:
          <cShortDescription>Ob du mit Gleichgesinnten über Fahrradthemen reden oder ob du nur ein Bier trinken willst:
            du bist willkommen beim Radl-Stammtisch im 'Bären' in Gauting.</cShortDescription>
    """
    return tour.getElementsByTagName("cShortDescription")[0].firstChild.data

def getSchwierigkeit():
    """
    Starnberg:
          <Schwierigkeit>***</Schwierigkeit>

     schwierigkeit = tour.getElementsByTagName("Schwierigkeit")[0].firstChild.data
    """
    kategorie = getKategorie()
    if kategorie == "Feierabendtour":
        return "F"
    if kategorie == "Halbtagestour" or kategorie == "Tagestour" or kategorie == "Mehrtagestour":
        return "??" # not yet in xml
    logger.error("unknown kategorie")
    return "???"

def getKategorie():
    itemTags = tour.getElementsByTagName("ItemTags")[0]
    logger.debug("itemTags %s %s ", type(itemTags), str(itemTags))
    expItemTags = itemTags.getElementsByTagName("ExportItemTag")
    logger.debug("expItemTags %s %s ", type(expItemTags), str(expItemTags))
    for exportItemTag in expItemTags:
        tag = exportItemTag.getElementsByTagName("Tag")[0].firstChild.data
        category = exportItemTag.getElementsByTagName("Category")[0].firstChild.data
        if category == "Typen (nach Dauer und Tageslage)":
            return tag
    return None

def getZusatzInfo():
    itemTags = tour.getElementsByTagName("ItemTags")[0]
    logger.debug("itemTags %s %s ", type(itemTags), str(itemTags))
    expItemTags = itemTags.getElementsByTagName("ExportItemTag")
    logger.debug("expItemTags %s %s ", type(expItemTags), str(expItemTags))
    besonders = ""
    weitere = ""
    zielgruppe = ""
    for exportItemTag in expItemTags:
        tag = exportItemTag.getElementsByTagName("Tag")[0].firstChild.data
        category = exportItemTag.getElementsByTagName("Category")[0].firstChild.data
        if category == "Besondere Charakteristik /Thema":
            besonders = besonders + tag + " "
        if category == "Weitere Eigenschaften":
            weitere = weitere + tag + " "
        if category == "Besondere Zielgruppe":
            zielgruppe = zielgruppe + tag + " "
    if besonders != "":
        besonders = "Besondere Charakteristik/Thema: " + besonders
    if weitere != "":
        weitere = "Weitere Eigenschaften: " + weitere
    if zielgruppe != "":
        zielgruppe = "Besondere Zielgruppe: " + zielgruppe
    return [besonders, weitere, zielgruppe]

def getStrecke():
    """
        Starnberg:
            <Strecke>10 km</Strecke>
        tour.getElementsByTagName("Strecke")[0].firstChild.data
    """
    return "?? km"

def getDatum():
    """
        Starnberg:
              <datum>Fr, 20.04.2018</datum>
        tour.getElementsByTagName("datum")[0].firstChild.data
        Neu:
             <Beginning>2018-07-06T14:00:00</Beginning>
    """
    datum = tour.getElementsByTagName("Beginning")[0].firstChild.data
    logger.debug("Beginning %s %s ", type(datum), str(datum))
    beginning = convertToMEZOrMSZ(datum)
    # fromisoformat defined in Python3.7, not used by Scribus
    # date = datetime.fromisoformat(datum)
    datum = str(datum[0:10])
    logger.debug("datum %s <%s> ", type(datum), str(datum))
    date = time.strptime(datum, "%Y-%m-%d")
    weekday = weekdays[date.tm_wday]
    res =  weekday + ", " + datum[8:10] + "." + datum[5:7] + "." + datum[0:4]
    logger.info("datum=%s", res)
    return res

def getEndDatum():
    """
        Starnberg:
              <enddatum>Fr, 20.04.2018</enddatum>
        tour.getElementsByTagName("datum")[0].firstChild.data
        Neu:
             <End>2018-07-06T14:00:00</End>
    """
    enddatums = tour.getElementsByTagName("End")
    # we see two <End> nodes!?
    logger.debug("enddatums %s %s ", type(enddatums), str(enddatums))
    for enddatum in enddatums:
        if enddatum.firstChild is None:
            continue
        enddatumData = enddatum.firstChild.data
        # fromisoformat defined in Python3.7, not used by Scribus
        # enddate = datetime.fromisoformat(enddatumData)
        enddatumData = str(enddatumData[0:10])
        #enddate = datetime.strptime(enddatumData, "%Y-%m-%d")
        enddate = time.strptime(enddatumData, "%Y-%m-%d")
        logger.debug("enddate %s %s ", type(enddate), str(enddate))
        weekday = weekdays[enddate.tm_wday]
        return weekday + ", " + enddatumData[8:10] + "." + enddatumData[5:7] + "." + enddatumData[0:4]
    return "???"

def getPersonen():
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
    """
    personen = []
    organizer = tour.getElementsByTagName("Organizer")
    if organizer is not None and len(organizer) > 0:
        organizer = organizer[0].firstChild.data
        personen.append(organizer)
    organizer2 = tour.getElementsByTagName("Organizer2")
    if organizer2 is not None and len(organizer2) > 0:
        organizer2 = organizer2[0].firstChild.data
        if organizer2 != organizer:
            personen.append(organizer2)
    if len(personen) == 0:
        logger.error("Tour %s hat keinen Tourleiter" )
    return personen

for tour in touren:
    try:
        titel = tour.getElementsByTagName("Title")[0].firstChild.data
        logger.info("tour %s %s %s", titel, type(tour), str(tour))

        kategorie = getKategorie()
        if kategorie is None:
            logger.error("Keine Radtour")
            continue    # keine Radtour

        abfahrten = getAbfahrten()
        if len(abfahrten) == 0:
            logger.error("kein Startpunkt in tour %s", titel)
            continue
        logger.info("abfahrten %s %s ", type(abfahrten), str(abfahrten))

        beschreibung = getBeschreibung()
        logger.info("beschreibung %s", beschreibung)
        zusatzinfo = getZusatzInfo()
        logger.info("zusatzinfo %s", str(zusatzinfo))
        logger.info("kategorie %s", kategorie)
        schwierigkeit = getSchwierigkeit()
        logger.info("schwierigkeit %s", schwierigkeit)
        strecke = getStrecke()
        logger.info("strecke %s", strecke)
        datum = getDatum()
        logger.info("datum %s", datum)

        if kategorie == 'Mehrtagestour':
            enddatum = getEndDatum()
            logger.info("enddatum %s", enddatum)

        personen = getPersonen()
        logger.info("personen %s", str(personen))
    except Exception as e:
        logger.error("error in tour with title %s: %s", titel, e)
        continue

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
