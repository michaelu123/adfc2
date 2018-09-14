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
import tourServer

keySta = "152085"
keyGau = "15208514"
keyMuc = "152059"
py2 = False

try:
    import scribusHandler
    import httplib  # scribus seems to use Python 2
    handler = scribusHandler.ScribusHandler()
    tourServerVar = tourServer.TourServer(True, handler.getUseRest())
    unitKey = handler.getUnitKey()
except ImportError:
    import printHandler
    import http.client as httplib
    import argparse
    parser = argparse.ArgumentParser(description="Formatiere Daten des Tourenportals")
    parser.add_argument("nummer", help="Gliederungsnummer, z.B. 152059 für München")
    parser.add_argument("useRest", help="Sollen aktuelle Daten vom Server geholt werden? (j/n)")
    args = parser.parse_args()
    unitKey = args.nummer
    yOrN = args.useRest.lower()[0]
    useRest = yOrN == 'j' or yOrN == 'y' or yOrN == 't'
    tourServerVar = tourServer.TourServer(False, useRest)
    handler = printHandler.PrintHandler()

touren = tourServerVar.getTouren(unitKey)
for tour in touren:
    eventItemId = tour.get("eventItemId");
    tour = tourServerVar.getTour(eventItemId)
    handler.handleTour(tour)
