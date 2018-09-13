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

import sys, time, locale
import xml.dom.minidom

try:
     import scribus
except ImportError:
     print "Unable to import the 'scribus' module. This script will only run within"
     print "the Python interpreter embedded in Scribus. Try Script->Execute Script."
     sys.exit(1)

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
    scribus.selectText(scribus.getTextLength()-n, n, textbox)
    scribus.deleteText(textbox)

def handleAbfahrt(abfahrt):
    uhrzeit = abfahrt.getElementsByTagName("Zeit")[0].firstChild.data
    ort = abfahrt.getElementsByTagName("Startort")[0].firstChild.data
    scribus.setStyle('Radtour_start',textbox)
    scribus.insertText('Start: '+uhrzeit+', '+ort+'\n',-1,textbox)

def handleTextfeld(stil,textelement):
    absaetze = textelement.getElementsByTagName("p")
    for absatz in absaetze:
        scribus.setStyle(stil,textbox)
        scribus.insertText(absatz.childNodes[0].data+'\n',-1,textbox)

def handleBeschreibung(textelement):
    handleTextfeld(textelement)

def handleTel(Name):
    telfestnetz = Name.getElementsByTagName("TelFestnetz")
    telmobil = Name.getElementsByTagName("TelMobil")
    if (len(telfestnetz)!=0):
        scribus.insertText(' ('+telfestnetz[0].firstChild.data+')',-1,textbox)
    if (len(telmobil)!=0):
        scribus.insertText(' ('+telmobil[0].firstChild.data+')',-1,textbox)

def handleName(TL):
    name = TL.getElementsByTagName("Name")[0]
    scribus.insertText(name.lastChild.data,-1,textbox)
    handleTel(name)

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
    scribus.setStyle('Radtour_titel',textbox)
    scribus.insertText(tt+'\n',-1,textbox)

def handleKopfzeile(dat, kat, schwierig, strecke):
    scribus.setStyle('Radtour_kopfzeile',textbox)
    scribus.insertText(dat+':	'+kat+'	'+schwierig+'	'+strecke+'\n',-1,textbox)

def handleKopfzeileMehrtage(anfang, ende, kat, schwierig, strecke):
    scribus.setStyle('Radtour_kopfzeile',textbox)
    scribus.insertText(anfang+' bis '+ende+':\n',-1,textbox)
    scribus.setStyle('Radtour_kopfzeile',textbox)
    scribus.insertText('	'+kat+'	'+schwierig+'	'+strecke+'\n',-1,textbox)

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

if (ftype != "TextFrame"):
     scribus.messageBox('Scribus - Script Error', "This is not a textframe. Try again.", scribus.ICON_WARNING, scribus.BUTTON_OK)
     sys.exit(2)

scribus.deleteText(textbox)
scribus.setStyle('Radtouren_titel',textbox)
scribus.insertText('Radtouren\n',0,textbox)

xmlfiledata = scribus.fileDialog('XML file', 'XML files(*.xml)')


DomTree = xml.dom.minidom.parse(xmlfiledata)
collection = DomTree.documentElement

# Alle Touren einlesen
touren = collection.getElementsByTagName("Termin")

# die Touren ausgeben

for tour in touren:

    abfahrten = tour.getElementsByTagName("Abfahrt")
    beschreibung = tour.getElementsByTagName("Beschreibung")[0]
    zusatzinfo = tour.getElementsByTagName("Zusatzinfo")[0]
    titel = tour.getElementsByTagName("TourTitel")[0].firstChild.data
    schwierigkeit = tour.getElementsByTagName("Schwierigkeit")[0].firstChild.data
    kategorie = tour.getElementsByTagName("Kategorie")[0].firstChild.data
    strecke = tour.getElementsByTagName("Strecke")[0].firstChild.data
    datum = tour.getElementsByTagName("datum")[0].firstChild.data
    #if len(tour.getElementsByTagName("enddatum")) > 0:
    if kategorie == 'Mehrtagestour':
        enddatum = tour.getElementsByTagName("enddatum")[0].firstChild.data
    personen = tour.getElementsByTagName("Leiter")[0].getElementsByTagName("Person")

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
    handleTextfeld('Radtour_zusatzinfo',zusatzinfo)

