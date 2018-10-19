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
import myLogger
import sys
import os
import time
import tourServer
import textHandler
import csvHandler

def toDate(dmy): # 21.09.2018
    d = dmy[0:2]
    m = dmy[3:5]
    if len(dmy) == 10:
        y = dmy[6:10]
    else:
        y = "20" + dmy[6:8]
    if y < "2018":
        raise ValueError("Kein Datum vor 2018 möglich")
    if int(d) == 0 or int(d) > 31 or int(m) == 0 or int(m) > 12 or int(y) < 2000 or int(y) > 2100:
        raise ValueError("Bitte Datum als dd.mm.jjjj angeben, nicht als " + dmy)
    return y + "-" + m + "-" + d # 2018-09-21

try:
    arg0 = sys.argv[0]
    if arg0.find(".py") == -1:
        raise ImportError
    else:
        import scribusHandler
    import httplib  # scribus seems to use Python 2
    handler = scribusHandler.ScribusHandler()
    tourServerVar = tourServer.TourServer(True, handler.getUseRest(), handler.getIncludeSub())
    unitKeys = handler.getUnitKeys().split(",")
    start = handler.getStart()
    end = handler.getEnd()
    type = handler.getType()
    radTyp = handler.getRad()
except ImportError:
    import printHandler
    import http.client as httplib
    import argparse
    parser = argparse.ArgumentParser(description="Formatiere Daten des Tourenportals")
    parser.add_argument("-a", "--aktuell", dest="useRest", action="store_true", help="Aktuelle Daten werden vom Server geholt")
    parser.add_argument("-u", "--unter", dest="includeSub", action="store_true", help="Untergliederungen einbeziehen")
    parser.add_argument("-f", "--format", dest="format", choices=["S", "M", "C"], help="Ausgabeformat (S=Starnberg, M=München, C=CSV", default="S")
    parser.add_argument("-t", "--type", dest="type", choices = ["R", "T", "A"], help="Typ (R=Radtour, T=Termin, A=alles), default=A", default="A")
    parser.add_argument("-r", "--rad", dest="radTyp", choices = ["R", "T", "M", "A"], help="Fahrradtyp (R=Rennrad, T=Tourenrad, M=Mountainbike, A=Alles), default=A", default="A")
    parser.add_argument("nummer", help="Gliederungsnummer(n), z.B. 152059 für München, komma-separierte Liste")
    parser.add_argument("start", help="Startdatum (TT.MM.YYYY)")
    parser.add_argument("end", help="Endedatum (TT.MM.YYYY)")
    args = parser.parse_args()
    unitKeys = args.nummer.split(",")
    useRest = args.useRest
    includeSub = args.includeSub
    start = args.start
    end = args.end
    type = args.type
    radTyp = args.radTyp
    tourServerVar = tourServer.TourServer(False, useRest, includeSub)
    format = args.format
    if format == "S":
        handler = printHandler.PrintHandler()
    elif format == "M":
        handler = textHandler.TextHandler()
    elif format == "C":
        handler = csvHandler.CsvHandler(sys.stdout)
    else:
        handler = printHandler.PrintHandler()

start = toDate(start)
end = toDate(end)

if type == "R":
    type = "Radtour"
elif type == "T":
    type = "Termin"
elif type == "A":
    type = "Alles"
else:
    raise ValueError("Typ muss R für Radtour, T für Termin, oder A für beides sein")

if radTyp == "R":
    radTyp = "Rennrad"
elif radTyp == "T":
    radTyp = "Tourenrad"
elif radTyp == "M":
    radTyp = "Mountainbike"
elif radTyp == "A":
    radTyp = "Alles"
else:
    raise ValueError("Rad muss R für Rennrad, T für Tourenrad, M für Mountainbike, oder A für alles sein")

touren = []
for unitKey in unitKeys:
    touren.extend(tourServerVar.getTouren(unitKey.strip(), start, end, type, isinstance(handler, textHandler.TextHandler)))

def tourdate(self):
    return self.get("beginning")
touren.sort(key=tourdate)  # sortieren nach Datum
if (isinstance(handler, textHandler.TextHandler) or isinstance(handler, csvHandler.CsvHandler)) and (type == "Radtour" or type == "Alles"):
    tourServerVar.calcNummern()

if len(touren) == 0:
    handler.nothingFound()
for tour in touren:
    tour = tourServerVar.getTour(tour)
    if tour.isTermin():
        handler.handleTermin(tour)
    else:
        if radTyp != "Alles" and tour.getRadTyp() != radTyp:
            continue
        handler.handleTour(tour)
