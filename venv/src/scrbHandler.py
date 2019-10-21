# encoding: utf-8

import tourRest
import selektion
import os
import sys
import re
import datetime
import time
#import markdown
from myLogger import logger

try:
    import scribus
except ImportError:
    print "Unable to import the 'scribus' module. This script will only run within"
    print "the Python interpreter embedded in Scribus. Try Script->Execute Script."
    sys.exit(1)

logger.debug("1scrb")
params = """
Linktyp: Frontend
Selektion:
	Gliederungen: 152085
	MitUntergliederungen: nein
	Beginn: 01.07.2019
	Ende: 31.07.2019
Terminselektion:
	Name: Stammtische
	Merkmalenthält: Stammtisch,Aktiventreff
	Name: Rest
	Merkmalenthältnicht: Stammtisch,Aktiventreff
Tourselektion:
	Name: Touren
	Radtyp: Tourenrad
	Name: MTB
	Radtyp: Mountainbike
	Name: RR
	Radtyp: Rennrad
	Name: MTT
	Kategorie: Mehrtagestour
""".split("\n")
logger.debug("2scrb %s", str(params))



schwierigkeitMap = { 0: "sehr einfach",
                     1: "sehr einfach",
                     2: "einfach",
                     3: "mittel",
                     4: "schwer",
                     5: "sehr schwer"}
# schwarzes Quadrat = Wingdings 2 0xA2, weißes Quadrat = 0xA3
schwierigkeitMMap = { 0: u"\u00a3\u00a3\u00a3\u00a3\u00a3",
                      1: u"\u00a2\u00a3\u00a3\u00a3\u00a3",
                      2: u"\u00a2\u00a2\u00a3\u00a3\u00a3",
                      3: u"\u00a2\u00a2\u00a2\u00a3\u00a3",
                      4: u"\u00a2\u00a2\u00a2\u00a2\u00a3",
                      5: u"\u00a2\u00a2\u00a2\u00a2\u00a2"}
paramRE = re.compile(r"(?u)\${(\w*?)}")
fmtRE = re.compile(r"(?u)\.fmt\((.*?)\)")
strokeRE = r'(\~{2})(.+?)\1'
ulRE = r'(\^{2})(.+?)\1'
STX = '\u0002'  # Use STX ("Start of text") for start-of-placeholder
ETX = '\u0003'  # Use ETX ("End of text") for end-of-placeholder
stxEtxRE = re.compile(r'%s(\d+)%s' % (STX, ETX))
headerFontSizes = [ 0, 24, 18, 14, 12, 10, 8 ] # h1-h6 headers have fontsizes 24-8
debug = False
nlctr = 0
adfc_blue = 0x004b7c  # CMYK=90 60 10 30
adfc_yellow = 0xee7c00 # CMYK=0 60 100 0
noPStyle = 'Default Paragraph Style'
noCStyle = 'Default Character Style'

logger.debug("3scrb")


def str2hex(s):
    return ":".join("{:04x}".format(ord(c)) for c in s)

"""
see https://stackoverflow.com/questions/4770297/convert-utc-datetime-string-to-local-datetime-with-python
"""
def convertToMEZOrMSZ(s):  # '2018-04-29T06:30:00+00:00'
    # py3: dt = time.strptime(s, "%Y-%m-%dT%H:%M:%S%z")
    dt = time.strptime(s[0:-6], "%Y-%m-%dT%H:%M:%S") # py2
    t = time.mktime(dt)
    dt1 = datetime.datetime.fromtimestamp(t)
    dt2 = datetime.datetime.utcfromtimestamp(t)
    diff = (dt1 - dt2).seconds
    t += diff
    dt = datetime.datetime.fromtimestamp(t)
    return dt

#  it seems that with "pyinstaller -F" tkinter (resp. TK) does not find data files relative to the MEIPASS dir
def pyinst(path):
    path = path.strip()
    if os.path.exists(path):
        return path
    if hasattr(sys, "_MEIPASS"): # i.e. if running as exe produced by pyinstaller
        pypath = sys._MEIPASS + "/" + path
        if os.path.exists(pypath):
            return pypath
    return path

def add_hyperlink(pos, url):
    pass

def insertHR():
    pass

"""
class ScrbTreeHandler(markdown.treeprocessors.Treeprocessor):
    def __init__(self, md):
        super().__init__(md)
        self.ancestors = []
        self.states = []
        self.nodeHandler = {
            "h1": self.h1,
            "h2": self.h2,
            "h3": self.h3,
            "h4": self.h4,
            "h5": self.h5,
            "h6": self.h6,
            "p": self.p,
            "strong": self.strong,
            "em": self.em,
            "blockquote": self.blockQuote,
            "stroke": self.stroke,
            "underline": self.underline,
            "ul": self.ul,
            "ol": self.ol,
            "li": self.li,
            "a": self.a,
            "img": self.img,
            "hr": self.hr }

    def setDeps(self, scrbHandler):
        self.scrbHandler = scrbHandler

    def unescape(self, m):
        return chr(int(m.group(1)))

    def printLines(self, s):
        s = stxEtxRE.sub(self.unescape, s)  # "STX40ETX" -> chr(40), see markdown/postprocessors/UnescapePostprocessor
        r = self.curPara.add_run(s) # style?
        r.bold = r.italic = r.font.strike = r.font.underline = False
        for fst in self.fontStyles:
            if fst == 'B':
                r.bold = True
            elif fst == 'I':
                r.italic = True
            elif fst == 'X':
                r.font.strike = True
            elif fst == 'U':
                r.font.underline = True
        self.curRun = r

    def walkOuter(self, node):
        global nlctr
        if debug:
            if node.text != None:
                ltext = node.text.replace("\n", "<" + str(nlctr) + "nl>")
                # node.text = node.text.replace("\n", str(nlctr) + "\n")
                nlctr += 1
            else:
                ltext = "None"
            if node.tail != None:
                ltail = node.tail.replace("\n", "<" + str(nlctr) + "nl>")
                # node.tail = node.tail.replace("\n", str(nlctr) + "\n")
                nlctr += 1
            else:
                ltail = "None"
            self.lvl += 4
            print(" " * self.lvl,"<<<<")
            print(" " * self.lvl, "node=", node.tag, ",text=", ltext,
                  "tail=", ltail)
        try:
            self.nodeHandler[node.tag](node)
            if node.tail is not None:
                self.printLines(node.tail)
        except Exception:
            msg = "Fehler während der Behandlung der Beschreibung des Events " +\
                  self.scrbHandler.tourMsg
            logger.exception(msg)
            print(msg)
        if debug:
            print(" " * self.lvl,">>>>")
            self.lvl -= 4

    def walkInner(self, node):
        if node.text is not None:
            self.printLines(node.text)
        for dnode in node:
            self.walkOuter(dnode)

    def h1(self, node):
        node.tail = None

    def h2(self, node):
        node.tail = None

    def h3(self, node):
        node.tail = None

    def h4(self, node):
        node.tail = None

    def h5(self, node):
        node.tail = None

    def h6(self, node):
        node.tail = None

    def p(self, node):
        node.tail = None

    def strong(self, node):
        pass

    def stroke(self, node):
        pass

    def underline(self, node):
        pass

    def em(self, node):
        pass

    def ul(self, node):
        node.text = node.tail = None
        self.walkInner(node)


    def ol(self, node):
        node.text = node.tail = None
        self.walkInner(node)

    def li(self, node):
        node.tail = None
        self.walkInner(node)

    def a(self, node):
        url = node.attrib["href"]
        pos = self.insertPos
        self.walkInner(node)
        add_hyperlink(pos, url)
        scribus.selectText(pos, self.insertPos - pos, self.textbox)
        scribus.setTextColor("ADFC_Yellow", self.textbox)

    def blockQuote(self, node):
        node.text = node.tail = None
        self.walkInner(node)

    def hr(self, node):
        node.tail = None
        insertHR()
        self.walkInner(node)

    def img(self, node):
        self.walkInner(node)

class ScrbExtension(markdown.Extension):
    def extendMarkdown(self, md):
        self.scrbTreeHandler = ScrbTreeHandler(md)
        md.treeprocessors.register(self.scrbTreeHandler, "scrbtreehandler", 5)
        md.inlinePatterns.register(
            markdown.inlinepatterns.SimpleTagInlineProcessor(
                strokeRE, 'stroke'), 'stroke', 40)
        md.inlinePatterns.register(
            markdown.inlinepatterns.SimpleTagInlineProcessor(
                ulRE, 'underline'), 'underline', 41)
"""

class ScrbRun:
    def __init__(self, text, style, charstyle, fontname, fontsize, textcolor):
        self.text = text
        self.style = style
        self.charstyle = charstyle
        self.fontname = fontname
        self.fontsize = fontsize
        self.textcolor = textcolor
    def __str__(self):
        return "\n{ text:" + self.text + ",pstyle:" + str(self.style) + ",cstyle:" + str(self.charstyle) + "}"
    def __repr__(self):
        return "\n{ text:" + self.text + ",type:" + str(type(self.text)) + ",pstyle:" + str(self.style) + ",cstyle:" + str(self.charstyle) + "}"

class ScrbHandler:
    def __init__(self, gui):
        logger.debug("1init")
        if not scribus.haveDoc():
            scribus.messageBox('Scribus - Script Error', "No document open", scribus.ICON_WARNING, scribus.BUTTON_OK)
            sys.exit(1)

        self.doc = None
        self.gui = gui
        self.terminselections = {}
        self.tourselections = {}
        self.touren = []
        self.termine = []
        self.url = None
        self.para = None
        self.run = None
        self.textbox = None
        global debug
        try:
            _ = os.environ["DEBUG"]
            debug = True
        except:
            debug = False
        # self.scrbExtension = ScrbExtension()
        # self.md = markdown.Markdown(extensions=[self.scrbExtension])
        # self.scrbExtension.scrbTreeHandler.setDeps(self)
        self.selFunctions = selektion.getSelFunctions()
        self.expFunctions = { # keys in lower case
            u"heute": self.expHeute,
            u"start": self.expStart,
            u"end": self.expEnd,
            u"nummer": self.expNummer,
            u"titel": self.expTitel,
            u"beschreibung": self.expBeschreibung,
            u"kurz": self.expKurzBeschreibung,
            u"tourleiter": self.expTourLeiter,
            u"betreuer": self.expBetreuer,
            u"name": self.expName,
            u"city": self.expCity,
            u"street": self.expStreet,
            u"kategorie": self.expKategorie,
            u"schwierigkeit": self.expSchwierigkeit,
            u"tourlänge": self.expTourLength,
            u"abfahrten": self.expAbfahrten,
            u"zusatzinfo": self.expZusatzInfo,
            u"höhenmeter": self.expHoehenMeter,
            u"character": self.expCharacter,
            u"schwierigkeitm": self.expSchwierigkeitM,
            u"abfahrtenm": self.expAbfahrtenM,
            u"tourleiterm": self.expTourLeiterM
        }
        logger.debug("2init")

    def openScrb(self, pp):
        logger.debug("1openScrb")
        self.parseParams(params)
        logger.debug("2openScrb")
        if pp:
            self.setGuiParams()
        logger.debug("3openScrb")
        allStyles = scribus.getAllStyles()
        paraStyles = scribus.getParagraphStyles()
        charStyles = scribus.getCharStyles()
        logger.debug("paraStyles: %s\ncharStyles:%s", str(paraStyles), str(charStyles))
        # wd2_style = doc_styles.add_style("WD2_STYLE", WD_STYLE_TYPE.CHARACTER)
        # wd2_font = wd2_style.font
        # wd2_font.name = "Wingdings 2"
        logger.debug("4openScrb")

    def nothingFound(self):
        logger.info("Nichts gefunden")
        print("Nichts gefunden")

    def makeRuns(self, pos1, pos2):
        runs = []
        scribus.selectText(pos1, 1, self.textbox)
        last_style = scribus.getStyle(self.textbox)
        last_charstyle = scribus.getCharacterStyle(self.textbox)
        last_fontname = scribus.getFont(self.textbox)
        last_fontsize = scribus.getFontSize(self.textbox)
        last_textcolor = scribus.getTextColor(self.textbox)

        text = u""
        changed = False
        for c in range(pos1, pos2):
            scribus.selectText(c, 1, self.textbox)
            char = scribus.getText(self.textbox)

            style = scribus.getStyle(self.textbox)
            if style != last_style:
                changed = True

            charstyle = scribus.getCharacterStyle(self.textbox)
            if charstyle != last_charstyle:
                changed = True

            fontname = scribus.getFont(self.textbox)
            if fontname != last_fontname:
                changed = True

            fontsize = scribus.getFontSize(self.textbox)
            if fontsize != last_fontsize:
                changed = True

            textcolor = scribus.getTextColor(self.textbox)
            if textcolor != last_textcolor:
                changed = True

            if changed:
                runs.append(ScrbRun(text, last_style, last_charstyle, last_fontname, last_fontsize, last_textcolor))
                text = u""
                last_fontname = fontname
                last_fontsize = fontsize
                last_textcolor = textcolor
                last_style = style
                last_charstyle = charstyle
                changed = False
            text = text + char
        return runs

    def insertText(self, text, run):
        pos = self.insertPos
        scribus.insertText(text, pos, self.textbox)
        tlen = len(unicode(text))
        logger.debug("insert pos=%d len=%d npos=%d", pos, tlen, pos+tlen)
        scribus.selectText(pos, tlen, self.textbox)
        scribus.setStyle(noPStyle if run.style is None else run.style, self.textbox)
        scribus.selectText(pos, tlen, self.textbox)
        scribus.setCharacterStyle(noCStyle if run.charstyle is None else run.charstyle, self.textbox)
        scribus.selectText(pos, tlen, self.textbox)
        scribus.setFont(run.fontname, self.textbox)
        scribus.selectText(pos, tlen, self.textbox)
        scribus.setFontSize(run.fontsize, self.textbox)
        scribus.selectText(pos, tlen, self.textbox)
        scribus.setTextColor(run.textcolor, self.textbox)
        scribus.selectText(pos, tlen, self.textbox)
        self.insertPos += tlen

    def parseParams(self, lines):
        texts = []
        #defaults:
        self.linkType = "Frontend"
        self.includeSub = True
        lx = 0
        selections = {}
        while lx < len(lines):
            line = lines[lx]
            words = line.split()
            if len(words) == 0:
                lx += 1
                continue
            word0 = words[0].lower().replace(":", "")
            if len(words) > 1:
                if word0 == "linktyp":
                    self.linkType = words[1].lower().capitalize()
                    lx += 1
                elif word0 == "ausgabedatei":
                    self.ausgabedatei = words[1]
                    lx += 1
                else:
                    raise ValueError(
                        "Unbekannter Parameter " + word0 +
                        ", erwarte linktyp oder ausgabedatei")
            elif word0 not in ["selektion", "terminselektion", "tourselektion" ]:
                raise ValueError(
                    "Unbekannter Parameter " + word0 +
                    ", erwarte selektion, terminselektion oder tourselektion")
            else:
                lx = self.parseSel(word0, lines, lx+1, selections)

        selection = selections.get("selektion")
        self.gliederung = selection.get("gliederungen")
        self.includeSub = selection.get("mituntergliederungen") == "ja"
        self.start = selection.get("beginn")
        self.end = selection.get("ende")

        sels = selections.get("terminselektion")
        if sels is not None:
            for sel in sels.values():
                self.terminselections[sel.get("name")] = sel
                for key in sel.keys():
                    if key != "name" and not isinstance(sel[key], list):
                        sel[key] = [ sel[key] ]

        sels = selections.get("tourselektion")
        if sels is not None:
            for sel in sels.values():
                self.tourselections[sel.get("name")] = sel
                for key in sel.keys():
                    if key != "name" and not isinstance(sel[key], list):
                        sel[key] = [ sel[key] ]

    def setGuiParams(self):
        if self.gui is None:
            return
        self.gui.setLinkType(self.linkType)
        if self.gliederung != None and self.gliederung != "":
            self.gui.setGliederung(self.gliederung)
        self.gui.setIncludeSub(self.includeSub)
        if self.start != None and self.start != "":
            self.gui.setStart(self.start)
        if self.end != None and self.end != "":
             self.gui.setEnd(self.end)
        self.setEventType()
        self.setRadTyp()

    def setEventType(self):
        typ = ""
        if len(self.terminselections) != 0 and len(self.tourselections) != 0:
            typ = "Alles"
        elif len(self.terminselections) != 0:
            typ = "Termin"
        elif len(self.tourselections) != 0:
            typ = "Radtour"
        if typ != "":
            self.gui.setEventType(typ)

    def setRadTyp(self):
        rts = set()
        for sel in self.tourselections.values():
            l = sel.get("radtyp")
            if l is None or len(l) == 0:
                l = [self.gui.getRadTyp()]
            for elem in l:
                rts.add(elem)
        if "Alles" in rts:
            typ = "Alles"
        elif len(rts) == 1:
            typ = rts.pop()
        else:
            typ = "Alles"
        self.gui.setRadTyp(typ)

    def getIncludeSub(self):
        return self.includeSub
    def getEventType(self):
        if len(self.terminselections) != 0 and len(self.tourselections) != 0:
            return "Alles"
        if len(self.terminselections) != 0:
            return "Termin"
        if len(self.tourselections) != 0:
            return "Radtour"
        return self.gui.getEventType()
    def getRadTyp(self):
        rts = set()
        for sel in self.tourselections.values():
            l = sel.get("radtyp")
            if l is None or len(l) == 0:
                l = [self.gui.getRadTyp()]
            for elem in l:
                rts.add(elem)
        if "Alles" in rts:
            return "Alles"
        if len(rts) == 1:
            return rts[0]
        return "Alles"
    def getUnitKeys(self):
        return self.gliederung
    def getStart(self):
        return self.start
    def getEnd(self):
        return self.end

    def parseSel(self, word, lines, lx, selections):
        selections[word] = sel = sel2 = {}
        while lx < len(lines):
            line = lines[lx]
            if line.strip() == "":
                lx += 1
                continue
            if not line[0].isspace():
                return lx
            words = line.split()
            word0 = words[0].lower().replace(":", "")
            if word0 == "name":
                word1 = words[1].lower()
                sel[word1] = sel2 = {}
                sel2["name"] = word1
            else:
                lst = ",".join(words[1:]).split(",")
                sel2[word0] = lst[0] if len(lst) == 1 else lst
            lx += 1
        return lx

    def handleTour(self, tour):
        self.touren.append(tour)

    def handleTermin(self, tour):
        self.termine.append(tour)

    def handleEnd(self):
        if self.doc is None:
            self.openScrb(False)
        """
        self.linkType = self.gui.getLinkType()
        self.gliederung = self.gui.getGliederung()
        self.includeSub = self.gui.getIncludeSub()
        self.start = self.gui.getStart()
        self.end = self.gui.getEnd()
        """
        self.linkType = "frontend"
        self.gliederung = "125085"
        self.includeSub = False
        self.start = "2019-07-01"
        self.end = "2019-07-31"

        pagenum = scribus.pageCount()
        for page in range(1, pagenum + 1):
            scribus.gotoPage(page)
            pageitems = scribus.getPageItems()
            for item in pageitems:
                self.textbox = item[0]
                if item[1] != 4:
                    continue
                textlen = scribus.getTextLength(self.textbox)
                scribus.selectText(0, textlen, self.textbox)
                logger.debug("textlen: " + str(textlen))
                if textlen == 0:
                    continue
                alltext = unicode(scribus.getAllText(self.textbox))
                self.evalTemplate(alltext)

        self.doc = None
        self.touren = []
        self.termine = []

    def evalPara(self, para):
        if debug:
            print("para", para.text)
        for run in para.runs:
            self.evalRun(run, None)

    def evalTemplate(self, alltext):
        logger.debug("alltext: %s %s", type(alltext), alltext)

        pos1 = alltext.find("/template")
        if pos1 < 0:
            return;
        pos2 = alltext.find("/endtemplate")
        if pos2 < 0:
            raise Exception("kein /endtemplate nach /template")
        pos2 += 12   # len("/endtemplate")
        lines = alltext[pos1:pos2].split('\r')
        logger.debug("lines:%s %s", type(lines), str(lines))
        line0 = lines[0]
        lineN = lines[-1]
        logger.debug("lineN: %s %s", type(lineN), lineN)
        if lineN != "/endtemplate":
            raise ValueError("Die letzte Zeile des templates darf nur /endtemplate enthalten")
        words = line0.split()
        typ = words[1]
        if typ != "/tour" and typ != "/termin":
            raise ValueError("Zweites Wort nach /template muß /tour oder /termin sein")
        typ = typ[1:]
        sel = words[2]
        if not sel.startswith("/selektion="):
            raise ValueError("Drittes Wort nach /template muß mit /selektion= beginnen")
        sel = sel[11:].lower()
        sels = self.tourselections if typ == "tour" else self.terminselections
        if not sel in sels:
            raise ValueError("Selektion " + sel + " nicht in " + typ + "selektion")
        sel = sels[sel]
        touren = self.touren if typ == "tour" else self.termine
        self.insertPos = pos1
        runs = self.makeRuns(pos1, pos2)
        logger.debug("runs:%s", str(runs))
        # can now remove template
        scribus.selectText(pos1, pos2-pos1, self.textbox)
        scribus.deleteText(self.textbox)
        self.insertPos = pos1
        self.evalTouren(sel, touren, runs)

    def evalTouren(self, sel, touren, runs):
        selectedTouren = []
        logger.debug("touren: %d", len(touren))
        for tour in touren:
            if selektion.selected(tour, sel):
                selectedTouren.append(tour)
        logger.debug("seltouren: %d", len(selectedTouren))
        if len(selectedTouren) == 0:
            return
        selectedTouren = selectedTouren[0:3]  # TODO weg
        for tour in selectedTouren:
            self.tourMsg = tour.getTitel() + " vom " + tour.getDatum()[0]
            logger.debug("tourMsg: %s", self.tourMsg)
            for run in runs:
                if run.text.lower().startswith("/kommentar"):
                    continue
                self.run = run
                rtext = run.text.strip()
                self.evalRun(tour)
                pos = self.insertPos
                if rtext == "${titel}" and self.url != None:
                    add_hyperlink(pos, self.url)

    def evalRun(self, tour):
        logger.debug("evalRun1 %s %s", type(self.run.text), self.run.text)
        linesOut = []
        linesIn = self.run.text.split('\r')
        for line in linesIn:
            if not line.startswith("/template") and\
                    not line.startswith("/endtemplate"):
                exp = self.expand(line, tour)
                if exp != None:
                    linesOut.append(exp)
        newtext = '\n'.join(linesOut)
        logger.debug("evalRun2 %s", newtext)
        self.insertText(newtext, self.run)

    def expand(self, s, tour):
        while True:
            mp = paramRE.search(s)
            if mp == None:
                logger.debug("noexp %s %s", type(s), s)
                return s
            gp = mp.group(1).lower()
            logger.debug("expand %s", gp)
            sp = mp.span()
            mf = fmtRE.search(s, pos=sp[1])
            if mf != None and sp[1] == mf.span()[0]: # i.e. if ${param] is followed immediately by .fmt()
                gf = mf.group(1)
                sf = mf.span()
                s = s[0:sf[0]] + s[sf[1]:]
                expanded = self.expandParam(gp, tour, gf)
            else:
                expanded = self.expandParam(gp, tour, None)
            if expanded == None: # special case for beschreibung, handled as markdown
                return None
            try:
                s = s[0:sp[0]] + expanded + s[sp[1]:]
            except Exception:
                logger.error("expanded = " + expanded)

    def expandParam(self, param, tour, format):
        try:
            f = self.expFunctions[param]
            return f(tour, format)
        except Exception as e:
            err = 'Fehler mit dem Parameter "' + param + \
                  '" des Events ' + self.tourMsg
            print(err)
            logger.exception(err)
            return param

    def expHeute(self, _, format):
        if format == None:
            return str(datetime.date.today())
        else:
            #return datetime.date.today().strftime(format)
            return datetime.datetime.now().strftime(format)

    def expStart(self, tour, format):
        dt = convertToMEZOrMSZ(tour.getDatumRaw())
        if format == None:
            return str(dt)
        else:
            return dt.strftime(format)

    def expEnd(self, tour, format):
        dt = convertToMEZOrMSZ(tour.getEndDatumRaw())
        if format == None:
            return str(dt)
        else:
            return dt.strftime(format)

    def expNummer(self, tour, _):
        k = tour.getKategorie()[0]
        if k == "T":
            k = "G" # Tagestour -> Ganztagestour
        return tour.getRadTyp()[0].upper() + " " + tour.getNummer() + " " + k

    def expTitel(self, tour, _):
        if self.linkType == "Frontend":
            self.url = tour.getFrontendLink()
        elif self.linkType == "Backend":
            self.url = tour.getBackendLink()
        else:
            self.url = None
        logger.info("Titel: " + tour.getTitel())
        return tour.getTitel()

    def expBeschreibung(self, tour, _):
        desc = tour.eventItem.get("description")
        desc = tourRest.removeSpcl(desc)
        desc = tourRest.removeHTML(desc)
        #desc = codecs.decode(desc, encoding = "unicode_escape")
        # self.md.convert(desc)
        # self.md.reset()
        # return None
        return desc

    def expName(self, tour, _):
        return tour.getName()
    def expKurzBeschreibung(self, tour, _):
        return tour.getShortDesc()
    def expCity(self, tour, _):
        return tour.getCity()
    def expStreet(self, tour, _):
        return tour.getStreet()

    def expKategorie(self, tour, _):
        return tour.getKategorie()

    def expSchwierigkeit(self, tour, _):
        return schwierigkeitMap[tour.getSchwierigkeit()]

    def expSchwierigkeitM(self, tour, _):
        self.run.fontname = "Wingdings 2 Regular"
        return schwierigkeitMMap[tour.getSchwierigkeit()]

    def expTourLength(self, tour, _):
        return tour.getStrecke()

    def expPersonen(self, bezeichnung, tour):
        tl = tour.getPersonen()
        if len(tl) == 0:
            return ""

        # print("TL0:", self.runX, "<<" + self.para.runs[self.runX].text + ">>", " ".join(["<" + run.text + ">" for run in self.para.runs]))
        run = self.para.add_run(text=bezeichnung + ": ", style=self.run.style)
        run.bold = True
        move_run_before(self.runX, self.para)
        # print("TL1:", " ".join(["<" + run.text + ">" for run in self.para.runs]))
        self.runX += 1

        self.para.add_run(text=", ".join(tl), style=self.run.style)
        move_run_before(self.runX, self.para)
        # print("TL2:", " ".join(["<" + run.text + ">" for run in self.para.runs]))
        return ""

    def expTourLeiter(self, tour, _):
        return self.expPersonen("Tourleiter", tour)

    def expBetreuer(self, tour, _):
        return self.expPersonen("Betreuer", tour)

    def expTourLeiterM(self, tour, _):
        tl = tour.getPersonen()
        if len(tl) == 0:
            return ""
        return ", ".join(tl)

    def expAbfahrten(self, tour, _):
        afs = tour.getAbfahrten()
        if len(afs) == 0:
            return ""
        afl = []
        for af in afs:
            if af[1] == "":
                afl.append(af[2])
            else:
                afl.append(af[0] + " " + af[1] + " " + af[2])
        #print("AB0:", self.runX, "<<" + self.para.runs[self.runX].text + ">>", " ".join(["<" + run.text + ">" for run in self.para.runs]))

        run = self.para.add_run(text="Ort" + ("" if len(afs) == 1 else "e") + ": ", style=self.run.style)
        run.bold = True
        move_run_before(self.runX, self.para)
        #print("AB1:", " ".join(["<" + run.text + ">" for run in self.para.runs]))
        self.runX += 1

        self.para.add_run(text=", ".join(afl), style=self.run.style)
        move_run_before(self.runX, self.para)
        #print("AB2:", " ".join(["<" + run.text + ">" for run in self.para.runs]))

        return ""

    def expZusatzInfo(self, tour, _):
        zi = tour.getZusatzInfo()
        if len(zi) == 0:
            return None
        for z in zi:
            #print("ZU0:", self.runX, "<<" + self.para.runs[self.runX].text + ">>",
            # " ".join(["<" + run.text + ">" for run in self.para.runs]))
            x = z.find(':') + 1
            run = self.para.add_run(text=z[0:x], style=self.run.style)
            run.bold = True
            move_run_before(self.runX, self.para)
            #print("ZU1:", " ".join(["<" + run.text + ">" for run in self.para.runs]))
            self.runX += 1

            self.para.add_run(text=z[x+1:] + "\n", style=self.run.style)
            move_run_before(self.runX, self.para)
            #print("ZU2:", " ".join(["<" + run.text + ">" for run in self.para.runs]))
            self.runX += 1
        return ""

    def expHoehenMeter(self, tour, _):
        return tour.getHoehenmeter()

    def expCharacter(self, tour, _):
        return tour.getCharacter()

    def expAbfahrtenM(self, tour, _):
        afs = tour.getAbfahrten()
        if len(afs) == 0:
            return ""
        s = afs[0][1] + " Uhr; " + afs[0][2]
        for afx, af in enumerate(afs[1:]):
            s = s + "\n " + str(afx+2) + ". Startpunkt: " + afs[afx][1] + " Uhr;" + afs[afx][2]
        return s

logger.debug("4scrb")
import tourServer
tourServerVar = tourServer.TourServer(True, False, False)
touren = tourServerVar.getTouren("152085", "2019-07-01", "2019-07-31", "Radtour")
sh = ScrbHandler(None)
for tour1 in touren:
    tour2 = tourServerVar.getTour(tour1)
    sh.handleTour(tour2)
sh.handleEnd()
logger.debug("6scrb")
