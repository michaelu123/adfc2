#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
import sys
import tourServer
import argparse

import vadb


def toDate(dmy):  # 21.09.2018
    d = dmy[0:2]
    m = dmy[3:5]
    if len(dmy) == 10:
        y = dmy[6:10]
    else:
        y = "20" + dmy[6:8]
    if y < "2018":
        raise ValueError("Kein Datum vor 2018 möglich")
    if int(d) == 0 or int(d) > 31 or int(m) == 0 or \
            int(m) > 12 or int(y) < 2000 or int(y) > 2100:
        raise ValueError("Bitte Datum als dd.mm.jjjj angeben, nicht als " + dmy)
    return y + "-" + m + "-" + d  # 2018-09-21

parser = argparse.ArgumentParser(description="Formatiere Daten des Tourenportals")
parser.add_argument("-a", "--aktuell", dest="useRest", action="store_true",
                    help="Aktuelle Daten werden vom Server geholt")
parser.add_argument("-u", "--unter", dest="includeSub", action="store_true",
                    help="Untergliederungen einbeziehen")
parser.add_argument("-t", "--type", dest="eventType", choices=["R", "T", "A"],
                    help="Typ (R=Radtour, T=Termin, A=Alles), default=A",
                    default="A")
parser.add_argument("-r", "--rad", dest="radTyp",
                    choices=["R", "T", "M", "A"],
                    help="Fahrradtyp (R=Rennrad, T=Tourenrad, M=Mountainbike, A=Alles), default=A",
                    default="A")
parser.add_argument("nummer",
                    help="Gliederungsnummer(n), z.B. 152059 für München, komma-separierte Liste")
parser.add_argument("start", help="Startdatum (TT.MM.YYYY)")
parser.add_argument("end", help="Endedatum (TT.MM.YYYY)")

# -u -t A -r A 182 01.08.2020 31.08.2020
args = parser.parse_args()
unitKeys = args.nummer.split(",")
useRest = args.useRest
includeSub = args.includeSub
start = args.start
end = args.end
eventType = args.eventType
radTyp = args.radTyp
tourServerVar = tourServer.EventServer(useRest, includeSub, 1)

start = toDate(start)
end = toDate(end)

if eventType == "R":
    eventType = "Radtour"
elif eventType == "T":
    eventType = "Termin"
elif eventType == "A":
    eventType = "Alles"
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
    raise ValueError("Rad muss R für Rennrad, T für Tourenrad, M für Mountainbike, oder A für Alles sein")

events = []
for unitKey in unitKeys:
    events.extend(tourServerVar.getEvents(unitKey.strip(), start, end, eventType))


events.sort(key=lambda x: x.get("beginning"))  # sortieren nach Datum
# tourServerVar.calcNummern()
with vadb.VADBHandler() as handler:
    for event in events:
        event = tourServerVar.getEvent(event)
        if event.isTermin():
            handler.handleTermin(event)
        else:
            if radTyp != "Alles" and event.getRadTyp() != radTyp:
                continue
            handler.handleTour(event)
            # break