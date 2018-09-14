# encoding: utf-8

from myLogger import logger

class DatenTest: # just log data
    def insertText(self, text):
        logger.info(text)
        print (text, end="")

class PrintHandler:
    def __init__(self):
        self.scribus = DatenTest()

    def handleAbfahrt(self, abfahrt):
        # abfahrt = (beginning, loc)
        uhrzeit = abfahrt[0]
        ort = abfahrt[1]
        logger.info("Abfahrt: uhrzeit=%s ort=%s", uhrzeit, ort)

    def handleTextfeld(self, stil,textelement):
        logger.info("Textfeld: stil=%s text=%s", stil, textelement)
        if textelement != None:
            self.scribus.insertText(textelement+'\n')

    def handleTextfeldList(self, stil,textList):
        logger.info("TextfeldList: stil=%s text=%s", stil, str(textList))
        for text in textList:
            if len(text) == 0:
                continue
            logger.info("Text: stil=%s text=%s", stil, text)
            self.scribus.insertText(text+'\n')

    def handleBeschreibung(self, textelement):
        self.handleTextfeld(textelement)

    def handleTel(self, Name):
        telfestnetz = Name.getElementsByTagName("TelFestnetz")
        telmobil = Name.getElementsByTagName("TelMobil")
        if len(telfestnetz)!=0:
            logger.info("Tel: festnetz=%s", telfestnetz[0].firstChild.data)
            self.scribus.insertText(' ('+telfestnetz[0].firstChild.data+')')
        if len(telmobil)!=0:
            logger.info("Tel: mobil=%s", telmobil[0].firstChild.data)
            self.scribus.insertText(' ('+telmobil[0].firstChild.data+')')

    def handleName(self, name):
        logger.info("Name: name=%s", name)
        self.scribus.insertText(name)
        # self.handleTel(name) ham wer nich!

    def handleTourenleiter(self, TLs):
        self.scribus.insertText('Tourenleiter: ')
        names = ", ".join(TLs)
        self.scribus.insertText(names)
        self.scribus.insertText('\n')

    def handleTitel(self, tt):
        logger.info("Titel: titel=%s", tt)
        self.scribus.insertText(tt+'\n')

    def handleKopfzeile(self, dat, kat, schwierig, strecke):
        logger.info("Kopfzeile: dat=%s kat=%s schwere=%s strecke=%s", dat, kat, schwierig, strecke)
        self.scribus.insertText(dat+':	'+kat+'	'+schwierig+'	'+strecke+'\n')

    def handleKopfzeileMehrtage(self, anfang, ende, kat, schwierig, strecke):
        logger.info("Mehrtage: anfang=%s ende=%s kat=%s schwere=%s strecke=%s", anfang, ende, kat, schwierig, strecke)
        self.scribus.insertText(anfang+' bis '+ende+':\n')
        self.scribus.insertText('	'+kat+'	'+schwierig+'	'+strecke+'\n')

    def handleTour(self, tour):
        try:
            titel = tour.getTitel()
            logger.info("Title %s", titel)
            datum = tour.getDatum()
            logger.info("datum %s", datum)

            abfahrten = tour.getAbfahrten()
            if len(abfahrten) == 0:
                raise ValueError("kein Startpunkt in tour %s", titel)
                return
            logger.info("abfahrten %s ", str(abfahrten))

            beschreibung = tour.getBeschreibung()
            logger.info("beschreibung %s", beschreibung)
            zusatzinfo = tour.getZusatzInfo()
            logger.info("zusatzinfo %s", str(zusatzinfo))
            kategorie = tour.getKategorie()
            logger.info("kategorie %s", kategorie)
            schwierigkeit = tour.getSchwierigkeit()
            logger.info("schwierigkeit %s", schwierigkeit)
            strecke = tour.getStrecke()
            logger.info("strecke %s", strecke)

            if kategorie == 'Mehrtagestour':
                enddatum = tour.getEndDatum()
                logger.info("enddatum %s", enddatum)

            personen = tour.getPersonen()
            logger.info("personen %s", str(personen))
        except Exception as e:
            logger.error("Fehler in der Tour %s: %s", titel, e)
            print("\nFehler in der Tour ", titel, ": ", e)
            return

        self.scribus.insertText('\n')
        if kategorie == 'Mehrtagestour':
            self.handleKopfzeileMehrtage(datum, enddatum, kategorie, schwierigkeit, strecke)
        else:
            self.handleKopfzeile(datum, kategorie, schwierigkeit, strecke)
            self.handleTitel(titel)
        for abfahrt in abfahrten:
            self.handleAbfahrt(abfahrt)
        self.handleTextfeld('Radtour_beschreibung',beschreibung)
        self.handleTourenleiter(personen)
        self.handleTextfeldList('Radtour_zusatzinfo',zusatzinfo)

