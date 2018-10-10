# encoding: utf-8

import tourRest
from myLogger import logger


class TextHandler:
    def nothingFound(self):
        logger.info("Nichts gefunden")
        print("Nichts gefunden")

    def handleTour(self, tour):
        try:
            titel = tour.getTitel()
            logger.info("Title %s", titel)
            tourNummer = tour.getNummer()
            radTyp = tour.getRadTyp()[0] # T,R,M
            tourTyp = tour.getKategorie()[0]
            if tourTyp == "T": # Tagestour
                tourTyp = "G"   # Ganztagstour...
            datum = tour.getDatum()
            logger.info("tourNummer %s radTyp %s tourTyp %s datum %s", tourNummer, radTyp, tourTyp, datum)

            abfahrten = tour.getAbfahrten()
            if len(abfahrten) == 0:
                raise ValueError("kein Startpunkt in tour %s", titel)
            logger.info("abfahrten %s ", str(abfahrten))

            beschreibung = tour.getBeschreibung(False)
            logger.info("beschreibung %s", beschreibung)
            zusatzinfo = tour.getZusatzInfo()
            logger.info("zusatzinfo %s", str(zusatzinfo))
            kategorie = tour.getKategorie()
            logger.info("kategorie %s", kategorie)
            schwierigkeit = str(tour.getSchwierigkeit())
            if schwierigkeit == "0":
                schwierigkeit = "1"
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

            personen = tour.getPersonen()
            logger.info("personen %s", str(personen))
            if len(personen) == 0:
                logger.error("Fehler: Tour %s hat keinen Tourleiter", titel)

        except Exception as e:
            logger.exception("Fehler in der Tour '%s': %s", titel, e)
            return

        print("{} ${} {} {}".format(titel, radTyp, tourNummer, tourTyp))
        print("{} ${}$ ${}".format(datum, strecke, schwierigkeit))
        if hoehenmeter != "0" and len(character) > 0:
            print("${} m; {}".format(hoehenmeter, character))
        elif hoehenmeter != "0":
            print("${} m".format(hoehenmeter))
        elif len(character) > 0:
            print(character)
        for abfahrt in abfahrten:
            print("${} Uhr; {}".format(abfahrt[0], abfahrt[1]))
        print(beschreibung)
        for info in zusatzinfo:
            if len(info) == 0:
                continue
            print(info)
        print("Leitung: {}".format(", ".join(personen)))

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

            beschreibung = tour.getBeschreibung(False)
            logger.info("beschreibung %s", beschreibung)
            zusatzinfo = tour.getZusatzInfo()
            logger.info("zusatzinfo %s", str(zusatzinfo))
            kategorie = tour.getKategorie()
            logger.info("kategorie %s", kategorie)

        except Exception as e:
            logger.exception("Fehler im Termin '%s': %s", titel, e)
            return

        print("{} - {}".format(titel, terminTyp)) # terminTyp z.B. Stammtisch, entbehrlich?
        print("{}".format(datum))
        for zeit in zeiten:
            print("${} Uhr; {}".format(zeit[0], zeit[1]))
        print(beschreibung)
        for info in zusatzinfo:
            if len(info) == 0:
                continue
            print(info)
        print()

