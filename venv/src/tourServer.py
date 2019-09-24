# encoding: utf-8
import os
import json
import tourRest
import functools
from myLogger import logger
import adfc_gliederungen

class TourServer:
    def __init__(self, py2, useResta, includeSuba):
        self.useRest = useResta
        self.includeSub = includeSuba
        self.tpConn = None
        self.alleTouren = []
        self.alleTermine = []
        self.py2 = False
        try:
            os.makedirs("c:/temp/tpjson")  # exist_ok = True does not work with Scribus (Python 2)
        except:
            pass
        if py2:
            import httplib  # scribus seems to use Python 2
            self.tpConn = httplib.HTTPSConnection("api-touren-termine.adfc.de")
            self.cacheMem = {}  # for Python2
            self.py2 = True
        else:
            import http.client as httplib
            self.tpConn = httplib.HTTPSConnection("api-touren-termine.adfc.de")
            self.cacheMem = None  # for Python3
            self.getUser = functools.lru_cache(maxsize=100)(self.getUser)
        self.loadUnits()

    def getTouren(self, unitKey, start, end, type):
        unit = "Alles" if unitKey is None or unitKey == "" else unitKey
        startYear = start[0:4]
        jsonPath = "c:/temp/tpjson/search-" + unit + ("_I_" if self.includeSub else "_") + startYear + ".json"
        if self.useRest or not os.path.exists(jsonPath):
            req = "/api/eventItems/search?limit=10000"
            par = ""
            if unitKey != None and unitKey != "":
                par += "&unitKey=" + unitKey
                if self.includeSub:
                    par += "&includeSubsidiary=true"
            par += "&beginning=" + startYear + "-01-01"
            par += "&end=" + startYear + "-12-31"
            req += par
            resp = self.httpget(req)
            if resp is None:
                return None
            jsRoot = json.load(resp)
        else:
            resp = None
            with open(jsonPath, "r") as jsonFile:
                jsRoot = json.load(jsonFile)
        items = jsRoot.get("items")
        touren = []
        if len(items) == 0:
            return touren
        if not resp is None:  # a REST call result always overwrites jsonPath
            with open(jsonPath, "w") as jsonFile:
                json.dump(jsRoot, jsonFile, indent=4)
        for item in iter(items):
            #item["imagePreview"] = ""  # save space
            titel = item.get("title")
            if titel is None:
                logger.error("Kein Titel für die Tour %s", str(item))
                continue
            if item.get("cStatus") == "Cancelled" or item.get("isCancelled"):
                logger.info("Tour %s ist gecancelt", titel)
                continue
            if type != "Alles" and item.get("eventType") != type:
                continue
            beginning = item.get("beginning")
            if beginning is None:
                logger.error("Kein Beginn für die Tour %s", titel)
                continue
            begDate = beginning[0:4]
            if begDate < start[0:4] or begDate > end[0:4]:
                continue
            if item.get("eventType") == "Radtour":
                self.alleTouren.append(item)
            else:
                self.alleTermine.append(item)
            begDate = tourRest.convertToMEZOrMSZ(beginning)[0:10]
            if begDate < start or begDate > end:
                logger.info("tour " + titel + " not in timerange")
                continue
            # add other filter conditions here
            touren.append(item)
        return touren

    def getTour(self, tourJsSearch):
        global tpConn
        eventItemId = tourJsSearch.get("eventItemId")
        imagePreview = tourJsSearch.get("imagePreview")
        escTitle = "".join([ (ch if ch.isalnum() else "_") for ch in tourJsSearch.get("title")])
        jsonPath = "c:/temp/tpjson/" + eventItemId[0:6] + "_" + escTitle + ".json"
        if self.useRest or not os.path.exists(jsonPath):
            resp = self.httpget("/api/eventItems/" + eventItemId)
            if resp is None:
                return None
            tourJS = json.load(resp)
            tourJS["eventItemFiles"] = None  # save space
            tourJS["images"] = []  # save space
            tourJS["imagePreview"] = imagePreview
            # if not os.path.exists(jsonPath):
            with open(jsonPath, "w") as jsonFile:
                json.dump(tourJS, jsonFile, indent=4)
        else:
            with open(jsonPath, "r") as jsonFile:
                tourJS = json.load(jsonFile)
        tour = tourRest.Tour(tourJS, tourJsSearch, self)
        return tour

    # not in py2 @functools.lru_cache(100)
    def getUser(self, userId):
        if self.cacheMem != None:
            val = self.cacheMem.get(userId)
            if val != None:
                return val
        global tpConn
        jsonPath = "c:/temp/tpjson/user_" + userId + ".json"
        if self.useRest or not os.path.exists(jsonPath):
            resp = self.httpget("/api/users/" + userId)
            if resp is None:
                return None
            else:
                userJS = json.load(resp)
                userJS["simpleEventItems"] = None
                # if not os.path.exists(jsonPath):
                with open(jsonPath, "w") as jsonFile:
                    json.dump(userJS, jsonFile, indent=4)
        else:
            with open(jsonPath, "r") as jsonFile:
                userJS = json.load(jsonFile)
        user = tourRest.User(userJS)
        if self.cacheMem != None:
            self.cacheMem[userId] = user
        return user

    def loadUnits(self):
        if self.py2:
            return
        global tpConn
        jsonPath = "c:/temp/tpjson/units.json"
        if not os.path.exists(jsonPath):
            resp = self.httpget("/api/units/")
            if resp is None:
                return None
            unitsJS = json.load(resp)
            with open(jsonPath, "w") as jsonFile:
                json.dump(unitsJS, jsonFile, indent=4)
        else:
            with open(jsonPath, "r", encoding="utf-8") as jsonFile:
                unitsJS = json.load(jsonFile)
        adfc_gliederungen.load(unitsJS)

    def calcNummern(self):
        self.alleTouren.sort(key=lambda x: x.get("beginning"))  # sortieren nach Datum
        yyyy = ""
        for tourJS in self.alleTouren:
            datum = tourJS.get("beginning")
            if datum[0:4] != yyyy:
                yyyy = datum[0:4]
                tnum = 100
                rnum = 300
                mnum = 400
                mtnum = 600
            tour = self.getTour(tourJS)
            radTyp = tour.getRadTyp()
            kategorie = tour.getKategorie()
            if kategorie == "Mehrtagestour":
                num = mtnum
                mtnum += 1
            elif radTyp == "Rennrad":
                num = rnum
                rnum += 1
            elif radTyp == "Mountainbike":
                num = mnum
                mnum += 1
            else:
                num = tnum
                tnum += 1
            tourJS["tourNummer"] = str(num)
        self.alleTermine.sort(key=lambda x: x.get("beginning"))  # sortieren nach Datum
        yyyy = ""
        for tourJS in self.alleTermine:
            datum = tourJS.get("beginning")
            if datum[0:4] != yyyy:
                yyyy = datum[0:4]
                tnum = 700
            num = tnum
            tnum += 1
            tourJS["tourNummer"] = str(num)

    def httpget(self, req):
        for retries in range(2):
            try:
                self.tpConn.request("GET", req)
            except Exception as e:
                logger.exception("error in request " + req)
                if isinstance(e, httplib.CannotSendRequest):
                    self.tpConn.close()
                    self.tpConn = httplib.HTTPSConnection("api-touren-termine.adfc.de")
                    continue
            try:
                resp = self.tpConn.getresponse()
            except Exception as e:
                logger.exception("cannot get response for " + req)
                if isinstance(e, httplib.ResponseNotReady):
                    self.tpConn.close()
                    self.tpConn = httplib.HTTPSConnection("api-touren-termine.adfc.de")
                    continue
            break

        try:
            if resp.status >= 300:
                logger.error("request %s failed: code %s reason %s: %s", req, resp.status, resp.reason, resp.read())
                return None
            else:
                logger.debug("resp %d %s", resp.status, resp.reason)
        except:
            pass
        return resp