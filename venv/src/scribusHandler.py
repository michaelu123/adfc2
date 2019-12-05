# encoding: utf-8
import sys

from myLogger import logger

try:
    import scribus
except ModuleNotFoundError:
    raise ImportError

yOrN = scribus.valueDialog("UseRest", "Sollen aktuelle Daten vom Server geholt werden? (j/n)").lower()[0]
useRest = yOrN == 'j' or yOrN == 'y' or yOrN == 't'
yOrN = scribus.valueDialog("IncludeSub", "Sollen Untergliederungen einbezogen werden? (j/n)").lower()[0]
includeSub = yOrN == 'j' or yOrN == 'y' or yOrN == 't'
eventType = scribus.valueDialog("Typ", "Typ (R=Radtour, T=Termin, A=Alles) (R/T/A)")
radTyp  = scribus.valueDialog("Fahrradtyp", "Fahrradtyp (R=Rennrad, T=Tourenrad, M=Mountainbike, A=Alles) (R/T/M/A)")
unitKeys = scribus.valueDialog("Gliederung(en)", "Bitte Nummer(n) der Gliederung angeben (komma-separiert)")
start = scribus.valueDialog("Startdatum", "Startdatum (TT.MM.YYYY)")
end = scribus.valueDialog("Endedatum", "Endedatum (TT.MM.YYYY)")


class ScribusHandler:
    def __init__(self):
        if not scribus.haveDoc():
            scribus.messageBox('Scribus - Script Error', "No document open", scribus.ICON_WARNING, scribus.BUTTON_OK)
            sys.exit(1)

        if scribus.selectionCount() == 0:
            scribus.messageBox('Scribus - Script Error',
                               "There is no object selected.\nPlease select a text frame and try again.",
                               scribus.ICON_WARNING, scribus.BUTTON_OK)
            sys.exit(2)
        if scribus.selectionCount() > 1:
            scribus.messageBox('Scribus - Script Error',
                               "You have more than one object selected.\nPlease select one text frame and try again.",
                               scribus.ICON_WARNING, scribus.BUTTON_OK)
            sys.exit(2)

        self.textbox = scribus.getSelectedObject()
        ftype = scribus.getObjectType(self.textbox)

        if ftype != "TextFrame":
            scribus.messageBox('Scribus - Script Error', "This is not a textframe. Try again.", scribus.ICON_WARNING,
                               scribus.BUTTON_OK)
            sys.exit(2)

        scribus.deleteText(self.textbox)
        scribus.setStyle('Radtouren_titel', self.textbox)
        scribus.insertText('Radtouren\n', 0, self.textbox)

    def getUseRest(self):
        return useRest
    def getIncludeSub(self):
        return includeSub
    def getUnitKeys(self):
        return unitKeys
    def getStart(self):
        return start
    def getEnd(self):
        return end
    def getEventType(self):
        return eventType
    def getRadTyp(self):
        return radTyp

    def addStyle(self, style, frame):
         try:
             scribus.setStyle(style, frame)
         except scribus.NotFoundError:
             scribus.createParagraphStyle(style)
             scribus.setStyle(style, frame)

    def nothingFound(self):
        logger.info("Nichts gefunden")
        scribus.insertText("Nichts gefunden\n", -1, self.textbox)

    def handleAbfahrt(self, abfahrt):
        # abfahrt = (type, beginning, loc)
        typ = abfahrt[0]
        uhrzeit = abfahrt[1]
        ort = abfahrt[2]
        logger.info("Abfahrt: type=%s uhrzeit=%s ort=%s", typ, uhrzeit, ort)
        scribus.setStyle('Radtour_start',self.textbox)
        scribus.insertText(typ + (': '+uhrzeit if uhrzeit != "" else "")+', '+ort+'\n', -1, self.textbox)

    def handleTextfeld(self, stil,textelement):
        logger.info("Textfeld: stil=%s text=%s", stil, textelement)
        if textelement != None:
            zeilen = textelement.split("\n")
            self.handleTextfeldList(stil, zeilen)

    def handleTextfeldList(self, stil, textList):
        logger.info("TextfeldList: stil=%s text=%s", stil, str(textList))
        for text in textList:
            if len(text) == 0:
                continue
            logger.info("Text: stil=%s text=%s", stil, text)
            scribus.setStyle(stil, self.textbox)
            scribus.insertText(text+'\n', -1, self.textbox)

    def handleTel(self, Name):
        telfestnetz = Name.getElementsByTagName("TelFestnetz")
        telmobil = Name.getElementsByTagName("TelMobil")
        if len(telfestnetz)!=0:
            logger.info("Tel: festnetz=%s", telfestnetz[0].firstChild.data)
            scribus.insertText(' ('+telfestnetz[0].firstChild.data+')', -1, self.textbox)
        if len(telmobil)!=0:
            logger.info("Tel: mobil=%s", telmobil[0].firstChild.data)
            scribus.insertText(' ('+telmobil[0].firstChild.data+')', -1, self.textbox)

    def handleName(self, name):
        logger.info("Name: name=%s", name)
        scribus.insertText(name, -1, self.textbox)
        # handleTel(name) ham wer nich!

    def handleTourenleiter(self, TLs):
        scribus.setStyle('Radtour_tourenleiter', self.textbox)
        scribus.insertText('Tourenleiter: ', -1, self.textbox)
        names = ", ".join(TLs)
        scribus.insertText(names, -1, self.textbox)
        scribus.insertText('\n', -1, self.textbox)

    def handleTitel(self, tt):
        logger.info("Titel: titel=%s", tt)
        scribus.setStyle('Radtour_titel', self.textbox)
        scribus.insertText(tt+'\n', -1, self.textbox)

    def handleKopfzeile(self, dat, kat, schwierig, strecke):
        logger.info("Kopfzeile: dat=%s kat=%s schwere=%s strecke=%s", dat, kat, schwierig, strecke)
        scribus.setStyle('Radtour_kopfzeile', self.textbox)
        scribus.insertText(dat+':	'+kat+'	'+schwierig+'	'+strecke+'\n', -1, self.textbox)

    def handleKopfzeileMehrtage(self, anfang, ende, kat, schwierig, strecke):
        logger.info("Mehrtage: anfang=%s ende=%s kat=%s schwere=%s strecke=%s", anfang, ende, kat, schwierig, strecke)
        scribus.setStyle('Radtour_kopfzeile', self.textbox)
        scribus.insertText(anfang+' bis '+ende+':\n',-1, self.textbox)
        scribus.setStyle('Radtour_kopfzeile', self.textbox)
        scribus.insertText('	'+kat+'	'+schwierig+'	'+strecke+'\n',-1, self.textbox)

    def handleTour(self, tour):
        try:
            titel = tour.getTitel()
            logger.info("Title %s", titel)
            datum = tour.getDatum()[0]
            logger.info("datum %s", datum)

            abfahrten = tour.getAbfahrten()
            if len(abfahrten) == 0:
                raise ValueError("kein Startpunkt in tour %s", titel)
            logger.info("abfahrten %s ", str(abfahrten))

            beschreibung = tour.getBeschreibung(True)
            logger.info("beschreibung %s", beschreibung)
            zusatzinfo = tour.getZusatzInfo()
            logger.info("zusatzinfo %s", str(zusatzinfo))
            kategorie = tour.getKategorie()
            radTyp = tour.getRadTyp()
            logger.info("kategorie %s radTyp %s", kategorie, radTyp)
            if kategorie == "Feierabendtour":
                schwierigkeit = "F"
            elif radTyp == "Rennrad":
                schwierigkeit = "RR"
            elif radTyp == "Mountainbike":
                schwierigkeit = "MTB"
            else:
                schwierigkeit = str(tour.getSchwierigkeit())
            if schwierigkeit == "0":
                schwierigkeit = "1"
            if schwierigkeit >= "1" and schwierigkeit <= "5":
                schwierigkeit = "*" * int(schwierigkeit)
            logger.info("schwierigkeit %s", schwierigkeit)
            strecke = tour.getStrecke()
            logger.info("strecke %s", strecke)

            if kategorie == 'Mehrtagestour':
                enddatum = tour.getEndDatum()[0]
                logger.info("enddatum %s", enddatum)

            personen = tour.getPersonen()
            logger.info("personen %s", str(personen))
            if len(personen) == 0:
                logger.error("Tour %s hat keinen Tourleiter", titel)
        except Exception as e:
            logger.exception("Fehler in der Tour %s: %s", titel, e)
            return

        scribus.insertText('\n', -1, self.textbox)
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

    def handleTermin(self, tour):
        pass
