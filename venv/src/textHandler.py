# encoding: utf-8

from myLogger import logger

class TextHandler:
    def handleTour(self, tour):
        try:
            titel = tour.getTitel()
            logger.info("Title %s", titel)
            sep = titel.rfind("#")
            if sep > 0:
                tourNummer = titel[sep+1:].strip()
                titel = titel[0:sep].strip()
            else:
                tourNummer = 999
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
            logger.info("schwierigkeit %s", schwierigkeit)
            strecke = tour.getStrecke()
            if strecke == "0 km":
                logger.error("Fehler: Tour %s hat keine Tourlänge", titel)
                print("Fehler: Tour {} hat keine Tourlänge".format(titel))
            logger.info("strecke %s", strecke)
            höhenmeter = tour.getHöhenmeter()
            character = tour.getCharacter()

            if kategorie == 'Mehrtagestour':
                enddatum = tour.getEndDatum()
                logger.info("enddatum %s", enddatum)

            personen = tour.getPersonen()
            logger.info("personen %s", str(personen))
            if len(personen) == 0:
                logger.error("Fehler: Tour %s hat keinen Tourleiter", titel)
                print("Fehler: Tour {} hat keinen Tourleiter".format(titel))

        except Exception as e:
            logger.error("Fehler in der Tour %s: %s", titel, e)
            print("Fehler in der Tour '", titel, "': ", e)
            return

        print("{} ${} {} {}".format(titel, radTyp, tourNummer, tourTyp))
        print("{} ${}$ ${}".format(datum, strecke, schwierigkeit))
        if höhenmeter != "0" and len(character) > 0:
            print("${} m; {}".format(höhenmeter, character))
        elif höhenmeter != "0":
            print("${} m".format(höhenmeter))
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
        print()

    def handleTermin(self, tour):
        try:
            titel = tour.getTitel()
            logger.info("Title %s", titel)
            tourTyp = tour.getKategorie()
            datum = tour.getDatum()
            logger.info("tourTyp %s datum %s", tourTyp, datum)

            abfahrten = tour.getAbfahrten()
            if len(abfahrten) == 0:
                raise ValueError("kein Startpunkt in tour %s", titel)
                return
            logger.info("abfahrten %s ", str(abfahrten))

            beschreibung = tour.getBeschreibung(False)
            logger.info("beschreibung %s", beschreibung)
            zusatzinfo = tour.getZusatzInfo()
            logger.info("zusatzinfo %s", str(zusatzinfo))
            kategorie = tour.getKategorie()
            logger.info("kategorie %s", kategorie)

        except Exception as e:
            logger.exception("Fehler in der Tour '%s': %s", titel, e)
            print("\nFehler in der Tour '", titel, "': ", e)
            return

        print("\n{} - {}".format(titel, tourTyp)) # tourTyp z.B. Stammtisch, entbehrlich?
        print("{}".format(datum))
        for abfahrt in abfahrten:
            print("${} Uhr; {}".format(abfahrt[0], abfahrt[1]))
        print(beschreibung)
        for info in zusatzinfo:
            if len(info) == 0:
                continue
            print(info)
        print()

"""
TODO: In Beschreibung ** und NL, <br>, <span>
"""
