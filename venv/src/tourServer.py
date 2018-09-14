# encoding: utf-8

import os
import json
import tourRest
from myLogger import logger

class TourServer:
    def __init__(self, py2, useResta):
        self.useRest = useResta
        self.tpConn = None
        if py2:
            import httplib  # scribus seems to use Python 2
            self.tpConn = httplib.HTTPSConnection("api-touren-termine.adfc.de")
        else:
            import http.client
            self.tpConn = http.client.HTTPSConnection("api-touren-termine.adfc.de")

    def getTouren(self, unitKey):
        jsonPath = "c:/temp/tpjson/search-" + unitKey + ".json"
        if self.useRest or not os.path.exists(jsonPath):
            self.tpConn.request("GET", "/api/eventItems/search?unitKey=" + unitKey)
            resp = self.tpConn.getresponse()
            logger.debug("resp %s %s ", type(resp), str(resp))
            jsRoot = json.load(resp)
        else:
            resp = None
            with open(jsonPath, "r") as jsonFile:
                jsRoot = json.load(jsonFile)

        items = jsRoot.get("items")
        touren = []
        for item in iter(items):
            item["imagePreview"] = ""  # save space
            titel = item.get("title")
            if titel is None:
                logger.error("Kein Titel für die Tour %s", str(item))
                continue;
            if item.get("eventType") != "Radtour":
                continue;
            if item.get("beginning") is None:
                logger.error("Kein Beginn für die Tour %s", str(item))
                continue;
            # add other filter conditions here
            touren.append(item)

        if not resp is None:  # and not os.path.exists(jsonPath):
            with open(jsonPath, "w") as jsonFile:
                json.dump(jsRoot, jsonFile, indent=4)

        def tourdate(self):
            return self.get("beginning")

        touren.sort(key=tourdate)  # sortieren nach Datum
        return touren

    def getTour(self, eventItemId):
        global tpConn
        jsonPath = "c:/temp/tpjson/" + eventItemId + ".json"
        if self.useRest or not os.path.exists(jsonPath):
            self.tpConn.request("GET", "/api/eventItems/" + eventItemId)
            resp = self.tpConn.getresponse()
            logger.debug("resp %d %s", resp.status, resp.reason)
            tourJS = json.load(resp)
            tourJS["eventItemFiles"] = None  # save space
            # if not os.path.exists(jsonPath):
            with open(jsonPath, "w") as jsonFile:
                json.dump(tourJS, jsonFile, indent=4)
        else:
            with open(jsonPath, "r") as jsonFile:
                tourJS = json.load(jsonFile)
        tour = tourRest.Tour(tourJS)
        return tour

