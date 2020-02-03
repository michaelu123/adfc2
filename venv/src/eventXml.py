import time
from xml.dom.minidom import *
import event
import tourServer
from myLogger import logger

schwierList = ["unbekannt", "sehr einfach", "einfach", "mittel", "schwer", "sehr schwer"]
eventServer = None

# from https://stackoverflow.com/questions/191536/converting-xml-to-json-using-python Paulo Vj
def parse_element(element):
    dict_data = dict()
    if element.nodeType == element.TEXT_NODE:
        dict_data['data'] = element.data
    if element.nodeType not in [element.TEXT_NODE, element.DOCUMENT_NODE,
                                element.DOCUMENT_TYPE_NODE]:
        for item in element.attributes.items():
            dict_data[item[0]] = item[1]
    if element.nodeType not in [element.TEXT_NODE, element.DOCUMENT_TYPE_NODE]:
        for child in element.childNodes:
            child_name, child_dict = parse_element(child)
            if child_name in dict_data:
                try:
                    dict_data[child_name].append(child_dict)
                except AttributeError:
                    dict_data[child_name] = [dict_data[child_name], child_dict]
            else:
                dict_data[child_name] = child_dict
    for k in dict_data.keys():
        v = dict_data[k]
        if isinstance(v, dict):
            if len(v) == 1 and "#text" in v.keys():
                dict_data[k] = v["#text"]["data"]
            elif len(v) == 0:
                dict_data[k] = ""
    return element.nodeName, dict_data


def elimText(d):
    if isinstance(d, dict):
        if '#text' in d:
            del d['#text']
        if 'xsi:nil' in d:
            del d['xsi:nil']
        for n in d:
            d[n] = elimText(d[n])
        if len(d) == 0:
            return ""
    elif isinstance(d, list):
        for n in d:
            elimText(n)
        pass
    return d

# need to remove non-XML chars from XML
class XMLFilter():
    def __init__(self, f):
        self.f = f

    def read(self, n):
        s = self.f.read(n)
        s = s.replace("&#x1", "§§§§")  # ????
        return s

def getTouren(fn):
    with open(fn, "r", encoding="utf-8") as f:
        f = XMLFilter(f)
        xmlt = parse(f)
    n, d = parse_element(xmlt)
    jsRoot = elimText(d)
    # jsonPath = "xml2.json"
    # with open(jsonPath, "w") as jsonFile:
    #     json.dump(jsRoot, jsonFile, indent=4)
    return jsRoot

class XmlEvent(event.Event):
    def __init__(self, eventItem):
        self.eventItem = eventItem
        self.tourLocations = eventItem.get("TourLocations")
        self.titel = self.eventItem.get("Title").strip()
        logger.info("eventItemId %s %s", self.titel, self.getEventItemId())

    def getTitel(self):
        return self.titel

    def getEventItemId(self):
        return self.eventItem.get("EventItemId")

    def getFrontendLink(self):
        evR = eventServer.getEvent({"eventItemId": self.getEventItemId(), "title": self.getTitel()})
        return evR.getFrontendLink()

    def getBackendLink(self):
        return "https://intern-touren-termine.adfc.de/modules/events/" + self.getEventItemId()

    def getNummer(self):
        num = self.eventJSSearch.get("eventNummer")
        if num is None:
            num = "999"
        return num

    def makeList(self, e):
        if not isinstance(e, list):
            return [e]
        return e

    def tourLoc(self, tl):
            if tl is None:
                return None
            typ = tl.get("Type")
            if typ != "Startpunkt" and typ != "Treffpunkt" and typ != "Zielort":
                return None
            beginning = tl.get("Beginning")
            logger.debug("beginning %s", beginning)  # '2018-04-24T12:00:00'
            beginning = event.convertToMEZOrMSZ(beginning)  # '2018-04-24T14:00:00'
            beginning = beginning[11:16]  # 14:00
            name = tl.get("Name")
            street = tl.get("Street")
            city = tl.get("City")
            logger.debug("name '%s' street '%s' city '%s'", name, street, city)
            loc = name
            if city != "":
                if loc == "":
                    loc = city
                else:
                    loc = loc + " " + city
            if street != "":
                if loc == "":
                    loc = street
                else:
                    loc = loc + " " + street
            if typ == "Startpunkt":
                if self.isTermin():
                    typ = "Treffpunkt"
                else:
                    typ = "Start"
            if typ == "Zielort":
                typ = "Ziel"
            return (typ, beginning, loc)

    def getAbfahrten(self):
        abfahrten = []
        tl = self.tourLoc(self.eventItem.get("CDeparture"))
        if tl is not None:
            abfahrten.append(tl)

        tls = self.eventItem.get("TourLocations")
        if tls != "":
            tls = tls.get("ExportTourLocation")
            tls = self.makeList(tls)
            for tl in tls:
                tl = self.tourLoc(tl)
                if tl is not None:
                    abfahrten.append(tl)

        tl = self.tourLoc(self.eventItem.get("CDestination"))
        if tl is not None:
            abfahrten.append(tl)
        return abfahrten

    def getBeschreibung(self, raw):
        desc = self.eventItem.get("Description")
        desc = event.removeHTML(desc)
        desc = event.removeSpcl(desc)
        if raw:
            return desc
        desc = event.normalizeText(desc)
        return desc

    def getKurzbeschreibung(self):
        desc = self.eventItem.get("CShortDescription")
        desc = event.normalizeText(desc)
        return desc

    def isTermin(self):
        return self.eventItem.get("EventType") == "Termin"

    def getSchwierigkeit(self):
        if self.isTermin():
            return "-"
        schwierigkeit = self.eventItem.get("CTourDifficulty")
        if schwierigkeit == "":
            return 0
        return schwierList.index(schwierigkeit)

    def getMerkmale(self):
        merkmale = []
        for itemTag in ["FurtherProperties", "UseableFor", "SpecialTargetGroup", "SpecialCharacteristic"]:
            tag = self.eventItem.get(itemTag)
            if tag is None or tag == "":
                continue
            tags = self.makeList(tag.get("ExportTag"))
            for tag in tags:
                merkmale.append(tag.get("Tag"))
        return merkmale

    """ 
    Tags: (LV BY Stand Feb 2020), d.h. keine Termin-Tags! 
    'FurtherProperties': 
        {'Badepause', 'Picknick (Selbstverpflegung)', 'Bahnfahrt', 'Einkehr in Restauration', 'Zusatzkosten ( z.B. Eintritte, Fährtickets)'}, 
    'UseableFor': 
        {'Pedelec', 'Rennrad', 'Tandem', 'Liegerad', 'Mountainbike', 'Alltagsrad'}, 
    'SpecialCharacteristic': 
        {'Natur', 'Neubürger-/Kieztouren', 'Kultur', 'Stadt entdecken / erleben'}, 
    'SpecialTargetGroup': 
        {'Familien', 'Touren für Kinder (bis 14 Jahren)', 'Senioren', 'Touren für Jugendliche (15-18 Jahren)', 'Jugendliche', 'Menschen mit Behinderungen'}
    """

    def getKategorie(self):
        # until portal issues fixed:
        evR = eventServer.getEvent({"eventItemId": self.getEventItemId(), "title": self.getTitel()})
        return evR.getKategorie()

    def getRadTyp(self):
        """
            Radtyp:
            "UseableFor": {
                "ExportTag": [
                    {
                        "Tag": "Alltagsrad"
                    },
                    {
                        "Tag": "Mountainbike"
                    },
                    {
                        "Tag": "Pedelec"
                    }
                ]
            },
        """
        # wenn nur Rennrad oder nur Mountainbike, dann dieses, sonst Tourenrad
        tag = self.eventItem.get("UseableFor")
        if tag is None or tag == "":
            return "Tourenrad"
        tags = self.makeList(tag.get("ExportTag"))
        rtCnt = len(tags)
        for tag in tags:
            t = tag.get("Tag")
            if rtCnt == 1 and (t == "Rennrad" or t == "Mountainbike"):
                return t
        return "Tourenrad"

    def getZusatzInfo(self):
        besonders = []
        weitere = []
        zielgruppe = []

        tag = self.eventItem.get("SpecialCharacteristic")
        if tag is not None and tag != "":
            besonders = [ x.get("Tag") for x in self.makeList(tag.get("ExportTag"))]

        tag = self.eventItem.get("FurtherProperties")
        if tag is not None and tag != "":
            weitere = [ x.get("Tag") for x in self.makeList(tag.get("ExportTag"))]

        tag = self.eventItem.get("SpecialTargetGroup")
        if tag is not None and tag != "":
            zielgruppe = [ x.get("Tag") for x in self.makeList(tag.get("ExportTag"))]

        zusatzinfo = []
        if len(besonders) > 0:
            besonders = "Besondere Charakteristik/Thema: " + ", ".join(besonders)
            zusatzinfo.append(besonders)
        if len(weitere) > 0:
            weitere = "Weitere Eigenschaften: " + ", ".join(weitere)
            zusatzinfo.append(weitere)
        if len(zielgruppe) > 0:
            zielgruppe = "Besondere Zielgruppe: " + ", ".join(zielgruppe)
            zusatzinfo.append(zielgruppe)
        return zusatzinfo

    def getStrecke(self):
        tl = self.eventItem.get("CTourLengthKm")
        return tl + " km"

    def getHoehenmeter(self):
        h = self.eventItem.get("CTourHeight")
        return h

    def getCharacter(self):
        c = self.eventItem.get("CTourSurface")
        return c

    def getDatum(self):
        """
            "Beginning": "2020-05-24T05:00:00",
            "BeginningDate": "24/05/2020",
            "BeginningTime": "05:00:00",
            "End": "2020-05-24T17:00:00",
            "EndDate": "24/05/2020",
            "EndTime": "05:00:00",
        """
        beginning = self.eventItem.get("Beginning")
        datum = event.convertToMEZOrMSZ(beginning)  # '2018-04-24T14:00:00'
        logger.debug("datum <%s>", str(datum))
        day = str(datum[0:10])
        date = time.strptime(day, "%Y-%m-%d")
        weekday = event.weekdays[date.tm_wday]
        res = (weekday + ", " + day[8:10] + "." + day[5:7] + "." + day[0:4], datum[11:16])
        return res

    def getDatumRaw(self):
        return self.eventItem.get("Beginning")

    def getEndDatum(self):
        beginning = self.eventItem.get("End")
        datum = event.convertToMEZOrMSZ(beginning)  # '2018-04-24T14:00:00'
        logger.debug("datum <%s>", str(datum))
        day = str(datum[0:10])
        date = time.strptime(day, "%Y-%m-%d")
        weekday = event.weekdays[date.tm_wday]
        res = (weekday + ", " + day[8:10] + "." + day[5:7] + "." + day[0:4], datum[11:16])
        return res

    def getEndDatumRaw(self):
        return self.eventItem.get("End")

    def getPersonen(self):
        personen = []
        org = self.eventItem.get("Organizer")
        if org is not None and org != "":
            personen.append(org)
        org2 = self.eventItem.get("Organizer2")
        if org2 is not None and org2 != "" and org2 != org:
            personen.append(str(org2))
        return personen

    def getImagePreview(self):
        return None

    def getName(self):
        dep = self.eventItem.get("CDeparture")
        return dep.get("Name")

    def getCity(self):
        dep = self.eventItem.get("CDeparture")
        return dep.get("City")

    def getStreet(self):
        dep = self.eventItem.get("CDeparture")
        return dep.get("Street")

    def isExternalEvent(self):
        return self.eventItem.get("CExternalEvent") == "Ja"

    def isEntwurf(self):
        return self.eventItem.get("CPublishDate") == ""


if __name__ == "__main__":
    eventServer = tourServer.EventServer(False, True, 1)
#    js = getTouren("C:\\Users\\Michael\\PycharmProjects\\ADFC1\\venv\\src\Veranstaltungs-Export.xml")
    js = getTouren("C:\\Users\\Michael\\Downloads\Veranstaltungs-Export_BY2020.xml")
    js = js.get("ExportEventItemList")
    js = js.get("EventItems")
    js = js.get("ExportEventItem")
    for ev in js:
        ev = XmlEvent(ev)
        x = ev.getTitel()
        print("Titel", x)
        x = ev.getEventItemId()
        print("Id", x)
        x = ev.getFrontendLink()
        print("Frontend", x)
        x = ev.getBackendLink()
        print("Backend", x)
        x = ev.getAbfahrten()
        print("Abfahrten", x)
        x = ev.getBeschreibung(False)
        print("Beschreibung", x)
        x = ev.getKurzbeschreibung()
        print("Kurzbes.", x)
        x = ev.isTermin()
        print("isTermin", x)
        x = ev.getSchwierigkeit()
        print("Schwierigkeit", x)
        x = ev.getMerkmale()
        print("Merkmale", x)
        x = ev.getKategorie()
        print("Kategorie", x)
        x = ev.getRadTyp()
        print("RadTyp", x)
        x = ev.getZusatzInfo()
        print("Zusatzinfo", x)
        x = ev.getStrecke()
        print("Strecke", x)
        x = ev.getHoehenmeter()
        print("Höhenmeter", x)
        x = ev.getCharacter()
        print("Character", x)
        x = ev.getDatum()
        print("Datum", x)
        x = ev.getDatumRaw()
        print("DatumRaw", x)
        x = ev.getEndDatum()
        print("EndDatum", x)
        x = ev.getEndDatumRaw()
        print("EndDatumRa", x)
        x = ev.getPersonen()
        print("Personen", x)
        x = ev.getImagePreview()
        #print(x)
        x = ev.getName()
        print("Name", x)
        x = ev.getCity()
        print("City", x)
        x = ev.getStreet()
        print("Street", x)
        x = ev.isExternalEvent()
        print("isExternalEvent", x)
        x = ev.isEntwurf()
        print("isEntwurf", x)
        print()
        print()
