# encoding: utf-8
import functools
import http.client
import json
import os
import threading
from concurrent.futures.thread import ThreadPoolExecutor

import adfc_gliederungen
import tourRest
from myLogger import logger


class EventServer:
    def __init__(self, useRest, includeSub, max_workers):
        self.useRest = useRest
        self.includeSub = includeSub
        self.max_workers = max_workers
        self.tpConns = []
        self.tpConnsLock = threading.Lock()
        self.events = {}
        self.alleTouren = []
        self.alleTermine = []
        self.py2 = False

        try:
            os.makedirs("c:/temp/tpjson")  # exist_ok = True does not work with Scribus (Python 2)
        except:
            pass
        self.getUser = functools.lru_cache(maxsize=100)(self.getUser)
        self.loadUnits()

    def getConn(self):
        with self.tpConnsLock:
            try:
                conn = self.tpConns.pop()
            except:
                conn = None
        if conn is None:
            conn = http.client.HTTPSConnection("api-touren-termine.adfc.de")
        return conn

    def putConn(self, conn):
        if conn is None:
            return
        with self.tpConnsLock:
            self.tpConns.insert(0, conn)

    def getEvents(self, unitKey, start, end, typ):
        unit = "Alles" if unitKey is None or unitKey == "" else unitKey
        startYear = start[0:4]
        jsonPath = "c:/temp/tpjson/search-" + unit + ("_I_" if self.includeSub else "_") + startYear + ".json"
        if self.useRest or not os.path.exists(jsonPath):
            req = "/api/eventItems/search?limit=10000"
            par = ""
            if unitKey is not None and unitKey != "":
                par += "&unitKey=" + unitKey
                if self.includeSub:
                    par += "&includeSubsidiary=true"
            par += "&beginning=" + startYear + "-01-01"
            par += "&end=" + startYear + "-12-31"
            req += par
            resp, conn = self.httpget(req)
            if resp is None:
                self.putConn(conn)
                return None
            jsRoot = json.load(resp)
            self.putConn(conn)
        else:
            resp = None
            with open(jsonPath, "r") as jsonFile:
                jsRoot = json.load(jsonFile)
        items = jsRoot.get("items")
        events = []
        if len(items) == 0:
            return events
        if resp is not None:  # a REST call result always overwrites jsonPath
            with open(jsonPath, "w") as jsonFile:
                json.dump(jsRoot, jsonFile, indent=4)
        for item in iter(items):
            # item["imagePreview"] = ""  # save space
            titel = item.get("title")
            if titel is None:
                logger.error("Kein Titel für den Event %s", str(item))
                continue
            if item.get("cStatus") == "Cancelled" or item.get("isCancelled"):
                logger.info("Event %s ist gecancelt", titel)
                continue
            if typ != "Alles" and item.get("eventType") != typ:
                continue
            beginning = item.get("beginning")
            if beginning is None:
                logger.error("Kein Beginn für den Event %s", titel)
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
                logger.info("event " + titel + " not in timerange")
                continue
            # add other filter conditions here
            logger.info("event " + titel + " OK")
            events.append(item)
        return events

    def getEvent(self, eventJsSearch):
        eventItemId = eventJsSearch.get("eventItemId")
        event = self.events.get(eventItemId)
        if event is not None:
            return event
        imagePreview = eventJsSearch.get("imagePreview")
        escTitle = "".join([(ch if ch.isalnum() else "_") for ch in eventJsSearch.get("title")])
        jsonPath = "c:/temp/tpjson/" + eventItemId[0:6] + "_" + escTitle + ".json"
        if self.useRest or not os.path.exists(jsonPath):
            resp, conn = self.httpget("/api/eventItems/" + eventItemId)
            if resp is None:
                self.putConn(conn)
                return None
            eventJS = json.load(resp)
            self.putConn(conn)
            eventJS["eventItemFiles"] = None  # save space
            eventJS["images"] = []  # save space
            eventJS["imagePreview"] = imagePreview
            # if not os.path.exists(jsonPath):
            with open(jsonPath, "w") as jsonFile:
                json.dump(eventJS, jsonFile, indent=4)
        else:
            with open(jsonPath, "r") as jsonFile:
                eventJS = json.load(jsonFile)
        event = tourRest.Event(eventJS, eventJsSearch, self)
        self.events[eventItemId] = event
        return event

    def getEventById(self, eventItemId, titel):
        ejs = {"eventItemId": eventItemId, "imagePreview": "", "title": titel}
        return self.getEvent(ejs)

    # not in py2 @functools.lru_cache(100)
    def getUser(self, userId):
        jsonPath = "c:/temp/tpjson/user_" + userId + ".json"
        if self.useRest or not os.path.exists(jsonPath):
            resp, conn = self.httpget("/api/users/" + userId)
            if resp is None:
                self.putConn(conn)
                return None
            userJS = json.load(resp)
            self.putConn(conn)
            userJS["simpleEventItems"] = None
            # if not os.path.exists(jsonPath):
            with open(jsonPath, "w") as jsonFile:
                json.dump(userJS, jsonFile, indent=4)
        else:
            with open(jsonPath, "r") as jsonFile:
                userJS = json.load(jsonFile)
        user = tourRest.User(userJS)
        return user

    def loadUnits(self):
        if self.py2:
            return
        jsonPath = "c:/temp/tpjson/units.json"
        if not os.path.exists(jsonPath):
            resp, conn = self.httpget("/api/units/")
            if resp is None:
                self.putConn(conn)
                return None
            unitsJS = json.load(resp)
            self.putConn(conn)
            with open(jsonPath, "w") as jsonFile:
                json.dump(unitsJS, jsonFile, indent=4)
        else:
            with open(jsonPath, "r", encoding="utf-8") as jsonFile:
                unitsJS = json.load(jsonFile)
        adfc_gliederungen.load(unitsJS)

    def calcNummern(self):
        # too bad we base numbers on kategorie and radtyp,which we cannot get from the search result
        ThreadPoolExecutor(max_workers=self.max_workers).map(self.getEvent, self.alleTouren)
        self.alleTouren.sort(key=lambda x: x.get("beginning"))  # sortieren nach Datum
        yyyy = ""
        logger.info("Begin calcNummern")
        for tourJS in self.alleTouren:
            datum = tourJS.get("beginning")
            if datum[0:4] != yyyy:
                yyyy = datum[0:4]
                tnum = 100
                rnum = 300
                mnum = 400
                mtnum = 600
            tour = self.getEvent(tourJS)
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
            tourJS["eventNummer"] = str(num)

        self.alleTermine.sort(key=lambda x: x.get("beginning"))  # sortieren nach Datum
        yyyy = ""
        for tourJS in self.alleTermine:
            datum = tourJS.get("beginning")
            if datum[0:4] != yyyy:
                yyyy = datum[0:4]
                tnum = 700
            num = tnum
            tnum += 1
            tourJS["eventNummer"] = str(num)
        logger.info("End calcNummern")

    def httpget(self, req):
        conn = None
        for retries in range(2):
            try:
                conn = self.getConn()
                conn.request("GET", req)
            except Exception as e:
                logger.exception("error in request " + req)
                if isinstance(e, http.client.CannotSendRequest):
                    conn.close()
                    conn = None
                    continue
            try:
                resp = conn.getresponse()
            except Exception as e:
                logger.exception("cannot get response for " + req)
                if isinstance(e, http.client.ResponseNotReady):
                    conn.close()
                    conn = None
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
        return (resp, conn)
