# encoding: utf-8

import tourRest
import csv
import sys
from myLogger import logger

schwierigkeitMap = { 0: "sehr einfach", 1: "sehr einfach", 2: "einfach", 3: "mittel", 4: "schwer", 5: "sehr schwer"}


class excel2(csv.Dialect):
    """Describe the usual properties of Excel-generated CSV files."""
    delimiter = ';'
    quotechar = '"'
    doublequote = True
    skipinitialspace = False
    lineterminator = '\n'
    quoting = csv.QUOTE_MINIMAL


class CsvHandler:
    def __init__(self, f):
        self.fieldNames = [ "Typ", "Titel", "Nummer", "Radtyp", "Tourtyp",
            "Datum", "Endedatum",
            "Tourlänge", "Schwierigkeit", "Höhenmeter", "Charakter",
            "Abfahrten", "Kurzbeschreibung", "Beschreibung", "ZusatzInfo", "Tourleiter"]

        csv.register_dialect("excel2", excel2)
        self.writer = csv.DictWriter(f, self.fieldNames, dialect="excel2")
        self.writer.writeheader()

    def nothingFound(self):
        logger.info("Nichts gefunden")
        print("Nichts gefunden")

    def handleTour(self, tour):
        try:
            titel = tour.getTitel()
            logger.info("Title %s", titel)
            tourNummer = tour.getNummer()
            radTyp = tour.getRadTyp()
            tourTyp = tour.getKategorie()
            datum = tour.getDatum()
            logger.info("tourNummer %s radTyp %s tourTyp %s datum %s", tourNummer, radTyp, tourTyp, datum)

            abfahrten = tour.getAbfahrten()
            if len(abfahrten) == 0:
                raise ValueError("kein Startpunkt in tour %s", titel)
            logger.info("abfahrten %s ", str(abfahrten))
            abfahrten = "\n".join([ "{}: {} Uhr; {}".format(abfahrt[0], abfahrt[1], abfahrt[2]) for abfahrt in abfahrten])
            beschreibung = tour.getBeschreibung(False)
            logger.info("beschreibung %s", beschreibung)
            kurzbeschreibung = tour.getKurzbeschreibung()
            logger.info("kurzbeschreibung %s", kurzbeschreibung)
            zusatzinfo = tour.getZusatzInfo()
            logger.info("zusatzinfo %s", str(zusatzinfo))
            zusatzinfo = "\n".join(zusatzinfo)
            kategorie = tour.getKategorie()
            logger.info("kategorie %s", kategorie)
            schwierigkeit = schwierigkeitMap[tour.getSchwierigkeit()]
            logger.info("schwierigkeit %s", schwierigkeit)
            strecke = tour.getStrecke()
            if strecke == "0 km":
                logger.error("Fehler: Tour %s hat keine Tourlänge", titel)
            else:
                logger.info("strecke %s", strecke)
            hoehenmeter = tour.getHoehenmeter()
            character = tour.getCharacter()

            if kategorie == 'Mehrtagestour':
                enddatum = tour.getEndDatum()
                logger.info("enddatum %s", enddatum)
            else:
                enddatum = ""

            tourLeiter = tour.getPersonen()
            logger.info("tourLeiter %s", str(tourLeiter))
            if len(tourLeiter) == 0:
                logger.error("Fehler: Tour %s hat keinen Tourleiter", titel)
            tourLeiter = ",".join(tourLeiter)

        except Exception as e:
            logger.exception("Fehler in der Tour '%s': %s", titel, e)
            return

        row = {
            "Typ":"Radtour", "Titel":titel, "Nummer":tourNummer, "Radtyp": radTyp, "Tourtyp": tourTyp,
            "Datum":datum, "Endedatum": enddatum,
            "Tourlänge": strecke, "Schwierigkeit": schwierigkeit, "Höhenmeter":hoehenmeter, "Charakter":character,
            "Abfahrten":abfahrten, "Kurzbeschreibung":kurzbeschreibung, "Beschreibung":beschreibung,
            "ZusatzInfo":zusatzinfo, "Tourleiter":tourLeiter }
        self.writer.writerow(row)

    def handleTermin(self, tour):
        try:
            titel = tour.getTitel()
            logger.info("Title %s", titel)
            terminTyp = tour.getKategorie()
            datum = tour.getDatum()
            logger.info("terminTyp %s datum %s", terminTyp, datum)

            zeiten = tour.getAbfahrten()
            if len(zeiten) == 0:
                raise ValueError("keine Anfangszeit für Termin %s", titel)
                return
            logger.info("zeiten %s ", str(zeiten))
            zeiten = "\n".join([ "{}: {} Uhr; {}".format(zeit[0], zeit[1], zeit[2]) for zeit in zeiten])
            beschreibung = tour.getBeschreibung(False)
            logger.info("beschreibung %s", beschreibung)
            kurzbeschreibung = tour.getKurzbeschreibung()
            logger.info("kurzbeschreibung %s", kurzbeschreibung)
            zusatzinfo = tour.getZusatzInfo()
            zusatzinfo = "\n".join(zusatzinfo)
            logger.info("zusatzinfo %s", str(zusatzinfo))

        except Exception as e:
            logger.exception("Fehler im Termin '%s': %s", titel, e)
            return

        row = {
            "Typ":"Termin", "Titel":titel, "Tourtyp": terminTyp,
            "Datum":datum,
            "Abfahrten":zeiten, "Kurzbeschreibung":kurzbeschreibung, "Beschreibung":beschreibung,
            "ZusatzInfo":zusatzinfo }
        self.writer.writerow(row)