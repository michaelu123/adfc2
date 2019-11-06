# encoding: utf-8

import copy
import datetime
import os
import re
import sys
import time

import adfc_gui
import markdown
import markdown.extensions.tables
import selektion
import styles
import tourRest
from myLogger import logger

try:
    import scribus
except ImportError:
    print("Unable to import the 'scribus' module. This script will only run within")
    print("the Python interpreter embedded in Scribus. Try Script->Execute Script.")
    sys.exit(1)

schwierigkeitMap = { 0: "sehr einfach",
                     1: "sehr einfach",
                     2: "einfach",
                     3: "mittel",
                     4: "schwer",
                     5: "sehr schwer"}
# schwarzes Quadrat = Wingdings 2 0xA2, weißes Quadrat = 0xA3
schwierigkeitMMap = { 0: "\u00a3\u00a3\u00a3\u00a3\u00a3",
                      1: "\u00a2\u00a3\u00a3\u00a3\u00a3",
                      2: "\u00a2\u00a2\u00a3\u00a3\u00a3",
                      3: "\u00a2\u00a2\u00a2\u00a3\u00a3",
                      4: "\u00a2\u00a2\u00a2\u00a2\u00a3",
                      5: "\u00a2\u00a2\u00a2\u00a2\u00a2"}
paramRE = re.compile(r"(?u)\${(\w*?)}")
fmtRE = re.compile(r"(?u)\.fmt\((.*?)\)")
strokeRE = r'(\~{2})(.+?)\1'
ulRE = r'(\^{2})(.+?)\1'
STX = '\u0002'  # Use STX ("Start of text") for start-of-placeholder
ETX = '\u0003'  # Use ETX ("End of text") for end-of-placeholder
stxEtxRE = re.compile(r'%s(\d+)%s' % (STX, ETX))
nlctr = 0
debug = False
adfc_blue = 0x004b7c  # CMYK=90 60 10 30
adfc_yellow = 0xee7c00 # CMYK=0 60 100 0
noPStyle = 'Default Paragraph Style'
noCStyle = 'Default Character Style'
lastPStyle = ""
lastCStyle = ""

def unicode(x): # because of Python2 compatibility
    return x

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

# add_hyperlink and insertHR can not be done in Scribus, as they are not related to the text, but to page
# coordinates. I.e. if the text changes, lines and boxes stay where they are. I cannot even draw a line after a line
# of text or draw a box around a word for a hyperlink, because I can't find out what the coordinates of the text are.
# see http://forums.scribus.net/index.php/topic,3487.0.html
def add_hyperlink(pos, url):
    pass

def insertHR():
    pass

class ScrbTreeProcessor(markdown.treeprocessors.Treeprocessor):
    def __init__(self, md):
        super().__init__(md)
        self.scrbHandler = None
        self.fontStyles = ""
        self.ctr = 0
        self.pIndent = 0
        self.numbered = False
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
            "hr": self.hr,
            "table": self.table,
            "thead": self.thead,
            "tbody": self.tbody,
            "tr": self.tr,
            "th": self.th,
            "td": self.td,
        }

    def run(self, root):
        self.tpRun = ScrbRun("", "MD_P_BLOCK", "MD_C_REGULAR")
        self.fontStyles = ""
        self.lvl = 4
        for child in root: # skip <div> root
            self.walkOuter(child)
        root.clear()

    def setDeps(self, scrbHandler):
        self.scrbHandler = scrbHandler

    def unescape(self, m):
        return chr(int(m.group(1)))

    def checkStylesExi(self, r):
        if not styles.checkCStyleExi(r.cstyle):
            r.cstyle = noCStyle
        if not styles.checkPStyleExi(r.pstyle):
            r.pstyle = noPStyle

    def printLines(self, s):
        s = stxEtxRE.sub(self.unescape, s)  # "STX40ETX" -> chr(40), see markdown/postprocessors/UnescapePostprocessor
        sav = self.tpRun.cstyle
        self.tpRun.cstyle = styles.modifyFont(self.tpRun.cstyle, self.fontStyles)
        self.scrbHandler.insertText(s, self.tpRun)
        self.tpRun.cstyle = sav

    def walkOuter(self, node):
        global nlctr
        if debug:
            if node.text is not None:
                ltext = node.text.replace("\n", "<" + str(nlctr) + "nl>")
                #node.text = node.text.replace("\n", str(nlctr) + "\n")
                nlctr += 1
            else:
                ltext = "None"
            if node.tail is not None:
                ltail = node.tail.replace("\n", "<" + str(nlctr) + "nl>")
                #node.tail = node.tail.replace("\n", str(nlctr) + "\n")
                nlctr += 1
            else:
                ltail = "None"
            self.lvl += 4
            logger.debug("MD:%s<<<<", " " * self.lvl)
            logger.debug("MD:%snode=%s,text=%s,tail=%s", " " * self.lvl, node.tag, ltext, ltail)
        try:
            self.nodeHandler[node.tag](node)
            if node.tail is not None:
                self.printLines(node.tail)
        except Exception:
            msg = "Fehler während der Behandlung der Beschreibung des Events " +\
                  self.scrbHandler.tourMsg
            logger.exception(msg)
        if debug:
            logger.debug("MD:%s>>>>", " " * self.lvl)
            self.lvl -= 4

    def walkInner(self, node):
        if node.tag == "li":
            node.text += node.tail
            node.tail = None
        if node.text is not None:
            self.printLines(node.text)
        for dnode in node:
                self.walkOuter(dnode)

    def h(self, pstyle, cstyle, node):
        savP = self.tpRun.pstyle
        savC = self.tpRun.cstyle
        self.tpRun.pstyle = pstyle
        self.tpRun.cstyle = cstyle
        self.checkStylesExi(self.tpRun)
        self.walkInner(node)
        self.tpRun.pstyle = savP
        self.tpRun.cstyle = savC

    def h1(self, node):
        self.h("MD_P_H1", "MD_C_H1", node)

    def h2(self, node):
        self.h("MD_P_H2", "MD_C_H2", node)

    def h3(self, node):
        self.h("MD_P_H3", "MD_C_H3", node)

    def h4(self, node):
        self.h("MD_P_H4", "MD_C_H4", node)

    def h5(self, node):
        self.h("MD_P_H5", "MD_C_H5", node)

    def h6(self, node):
        self.h("MD_P_H6", "MD_C_H6", node)

    def p(self, node):
        self.tpRun.cstyle = "MD_C_REGULAR"
        self.checkStylesExi(self.tpRun)
        self.walkInner(node)

    def strong(self, node):
        sav = self.fontStyles
        self.fontStyles += "B"
        self.walkInner(node)
        self.fontStyles = sav

    def stroke(self, node):
        sav = self.fontStyles
        self.fontStyles += "X"
        self.walkInner(node)
        self.fontStyles = sav

    def underline(self, node):
        sav = self.fontStyles
        self.fontStyles += "U"
        self.walkInner(node)
        self.fontStyles = sav

    def em(self, node):
        sav = self.fontStyles
        self.fontStyles += "I"
        self.walkInner(node)
        self.fontStyles = sav

    def plist(self, node, numbered):
        node.text = node.tail = None
        savCtr = self.ctr
        savP = self.tpRun.pstyle
        savN = self.numbered
        self.ctr = 1
        self.pIndent += 1
        self.numbered = numbered
        self.tpRun.pstyle = styles.listStyle(savP, numbered, self.ctr, self.pIndent)
        self.walkInner(node)
        self.pIndent -= 1
        self.ctr = savCtr
        self.tpRun.pstyle = savP
        self.numbered = savN

    def ul(self, node): # bullet
        self.plist(node, False)

    def ol(self, node): # numbered
        self.plist(node, True)

    def li(self, node):
        if self.numbered:
            self.scrbHandler.insertText(str(self.ctr)  + ".\t", self.tpRun)
        else:
            savC = self.tpRun.cstyle
            self.tpRun.cstyle = styles.bulletStyle()
            self.scrbHandler.insertText(styles.BULLET_CHAR + "\t", self.tpRun)
            self.tpRun.cstyle = savC
        self.ctr += 1
        self.walkInner(node)

    def a(self, node):
        url = node.attrib["href"]
        #pos = self.insertPos
        self.walkInner(node)
        #add_hyperlink(pos, url)
        #scribus.selectText(pos, self.insertPos - pos, self.textbox)
        #scribus.setTextColor("ADFC_Yellow", self.textbox)

    def blockQuote(self, node):
        node.text = node.tail = None
        savP = self.tpRun.pstyle
        self.tpRun.pstyle = "MD_P_BLOCK"
        self.checkStylesExi(self.tpRun)
        self.walkInner(node)
        self.tpRun.pstyle = savP

    def hr(self, node):
        node.tail = None
        insertHR()
        self.walkInner(node)

    def img(self, node):
        self.walkInner(node)

    def table(self, node):
        node.text = node.tail = None
        savP = self.tpRun.pstyle
        self.tpRun.pstyle = "MD_P_REGULAR"
        self.checkStylesExi(self.tpRun)
        self.walkInner(node)
        self.tpRun.pstyle = savP

    def thead(self, node):
        pass

    def tbody(self, node):
        node.text = node.tail = ""
        self.walkInner(node)

    def tr(self, node):
        node.text = ""
        l = len(node) - 1
        # separate td's by \t, that's all I can do at the moment
        for i,dnode in enumerate(node):
            if i != l:
                dnode.text += "\t"
        self.walkInner(node)

    def th(self, node):
        pass

    def td(self, node):
        node.tail = ""
        self.walkInner(node)

class ScrbExtension(markdown.Extension):
    def __init__(self):
        super(ScrbExtension,self).__init__()
        self.scrbTreeProcessor = None

    def extendMarkdown(self, md, globals):
        self.scrbTreeProcessor = ScrbTreeProcessor(md)
        md.treeprocessors.register(self.scrbTreeProcessor, "scrbTreeProcessor", 5)
        md.inlinePatterns.register(
            markdown.inlinepatterns.SimpleTagInlineProcessor(
                strokeRE, 'stroke'), 'stroke', 40)
        md.inlinePatterns.register(
            markdown.inlinepatterns.SimpleTagInlineProcessor(
                ulRE, 'underline'), 'underline', 41)


class ScrbRun:
    def __init__(self, text, pstyle, cstyle):
        self.text = text
        self.pstyle = pstyle
        self.cstyle = cstyle
    def __str__(self):
        return "\n{ text:" + self.text + ",type:" + str(type(self.text)) + ",pstyle:" + str(self.pstyle) + ",cstyle:" + str(self.cstyle) + "}"
    def __repr__(self):
        return "\n{ text:" + self.text + ",type:" + str(type(self.text)) + ",pstyle:" + str(self.pstyle) + ",cstyle:" + str(self.cstyle) + "}"

class ScrbHandler:
    def __init__(self, gui):
        if not scribus.haveDoc():
            scribus.messageBox('Scribus - Script Error', "No document open", scribus.ICON_WARNING, scribus.BUTTON_OK)
            sys.exit(1)

        self.gui = gui
        self.terminselections = {}
        self.tourselections = {}
        self.touren = []
        self.termine = []
        self.url = None
        self.run = None
        self.textbox = None
        self.linkType = "Frontend"
        self.gliederung = None
        self.includeSub = False
        self.start = None
        self.end = None
        self.pos = None
        self.ausgabedatei = None

        global debug
        try:
            _ = os.environ["DEBUG"]
            debug = True
        except:
            debug = True #False

        self.openScrb()
        self.scrbExtension = ScrbExtension()
        self.md = markdown.Markdown(extensions=[self.scrbExtension, markdown.extensions.tables.makeExtension()], enable_attributes=True, logger=logger)
        self.scrbExtension.scrbTreeProcessor.setDeps(self)
        self.selFunctions = selektion.getSelFunctions()
        self.expFunctions = { # keys in lower case
            "heute": self.expHeute,
            "start": self.expStart,
            "end": self.expEnd,
            "nummer": self.expNummer,
            "titel": self.expTitel,
            "beschreibung": self.expBeschreibung,
            "kurz": self.expKurzBeschreibung,
            "tourleiter": self.expTourLeiter,
            "betreuer": self.expBetreuer,
            "name": self.expName,
            "city": self.expCity,
            "street": self.expStreet,
            "kategorie": self.expKategorie,
            "schwierigkeit": self.expSchwierigkeit,
            "tourlänge": self.expTourLength,
            "abfahrten": self.expAbfahrten,
            "zusatzinfo": self.expZusatzInfo,
            "höhenmeter": self.expHoehenMeter,
            "character": self.expCharacter,
            "schwierigkeitm": self.expSchwierigkeitM,
            "abfahrtenm": self.expAbfahrtenM,
            "tourleiterm": self.expTourLeiterM
        }

    def openScrb(self):
        paraStyles = scribus.getParagraphStyles()
        charStyles = scribus.getCharStyles()
        logger.debug("paraStyles: %s\ncharStyles:%s", str(paraStyles), str(charStyles))
        self.parseParams()
        scribus.defineColor("ADFC_Yellow_", 0, 153, 255, 0)
        scribus.defineColor("ADFC_Blue_", 230, 153, 26, 77)
        scribus.createCharStyle(name="WD2_Y_", font="Wingdings 2 Regular", fillcolor = "ADFC_Yellow_")

    def nothingFound(self):
        logger.info("Nichts gefunden")
        print("Nichts gefunden")

    def makeRuns(self, pos1, pos2):
        runs = []

        scribus.selectText(pos1, pos2 - pos1, self.textbox)
        txtAll = scribus.getAllText(self.textbox)

        scribus.selectText(pos1, 1, self.textbox)
        last_pstyle = scribus.getStyle(self.textbox)
        last_cstyle = scribus.getCharacterStyle(self.textbox)

        text = ""
        changed = False
        for c in range(pos1, pos2):
            scribus.selectText(c, 1, self.textbox)
            # does not work reliably, see https://bugs.scribus.net/view.php?id=15911
            # char = scribus.getText(self.textbox)
            char = txtAll[c-pos1]

            pstyle = scribus.getStyle(self.textbox)
            if pstyle != last_pstyle:
                changed = True

            cstyle = scribus.getCharacterStyle(self.textbox)
            if cstyle != last_cstyle:
                changed = True

            # ff = scribus.getFontFeatures(self.textbox)
            # if ff != last_ff:
            #     # ff mostly "", for Wingdins chars ="-clig,-liga" !?!?
            #     logger.debug("fontfeature %s", ff)
            #     last_ff = ff

            if changed:
                runs.append(ScrbRun(text, last_pstyle, last_cstyle))
                last_pstyle = pstyle
                last_cstyle = cstyle
                text = ""
                changed = False
            text = text + char
        if text != "":
            runs.append(ScrbRun(text, last_pstyle, last_cstyle))
        return runs

    def insertText(self, text, run):
        if text is None or text == "":
            return
        pos = self.insertPos
        scribus.insertText(text, pos, self.textbox)
        tlen = len(unicode(text))
        logger.debug("insert pos=%d len=%d npos=%d text='%s' style=%s cstyle=%s",
                     pos, tlen, pos+tlen, text, run.pstyle, run.cstyle)
        global lastPStyle, lastCStyle
        if run.pstyle != lastPStyle:
            scribus.selectText(pos, tlen, self.textbox)
            scribus.setStyle(noPStyle if run.pstyle is None else run.pstyle, self.textbox)
            lastPStyle = run.pstyle
        if run.cstyle != lastCStyle:
            scribus.selectText(pos, tlen, self.textbox)
            scribus.setCharacterStyle(noCStyle if run.cstyle is None else run.cstyle, self.textbox)
            lastCStyle = run.cstyle
        if False and self.url is not None: # TODO
            logger.debug("URL: %s", self.url)
            scribus.selectText(pos, tlen, self.textbox)
            frame = None # see http://forums.scribus.net/index.php/topic,3487.0.html
            scribus.setURIAnnotation(self.url, frame)
            self.url = None
        self.insertPos += tlen
        # if scribus.textOverflows(self.textbox):
        #    self.createNewPage()

    def createNewPage(self):
        curPage = scribus.currentPage()
        if  curPage < scribus.pageCount() - 1:
            where = curPage + 1
        else:
            where = -1
        logger.debug("cur=%d pc=%d wh=%d", curPage, scribus.pageCount(), where)
        cols = scribus.getColumns(self.textbox)
        colgap = scribus.getColumnGap(self.textbox)
        x,y = scribus.getPosition(self.textbox)
        w,h = scribus.getSize(self.textbox)
        mp = scribus.getMasterPage(curPage)
        scribus.newPage(where, mp) # return val?
        scribus.gotoPage(curPage + 1)
        newFrame = scribus.createText(x,y,w,h)
        scribus.setColumns(cols, newFrame)
        scribus.setColumnGap(colgap, newFrame)
        scribus.linkTextFrames(self.textbox, newFrame)
        self.textbox = newFrame
        self.insertPos = 0 # scribus.getTextLength(newFrame)

    def parseParams(self):
        pagenum = scribus.pageCount()
        for page in range(1, pagenum + 1):
            scribus.gotoPage(page)
            lines = []
            pageitems = scribus.getPageItems()
            for item in pageitems:
                if item[1] != 4:
                    continue
                self.textbox = item[0]
                textlen = scribus.getTextLength(self.textbox)
                if textlen == 0:
                    continue
                scribus.selectText(0, textlen, self.textbox)
                alltext = unicode(scribus.getAllText(self.textbox))
                pos1 = alltext.find("/parameter")
                if pos1 < 0:
                    continue
                pos2 = alltext.find("/endparameter")
                if pos2 < 0:
                    raise ValueError("kein /endparameter nach /parameter")
                pos2 += 13 # len("/endparameter")
                lines = alltext[pos1:pos2].split('\r')[1:-1]
                logger.debug("parsePar lines:%s %s", type(lines), str(lines))
                scribus.selectText(pos1, pos2 - pos1, self.textbox)
                scribus.deleteText(self.textbox)
                break

        if len(lines) == 0:
            raise ValueError("No /parameter - /endparameter section in document")
        logger.debug("9parsePar")
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
        if self.gliederung is not None and self.gliederung != "":
            self.gui.setGliederung(self.gliederung)
        self.gui.setIncludeSub(self.includeSub)
        if self.start is not None and self.start != "":
            self.gui.setStart(self.start)
        if self.end is not None and self.end != "":
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
        self.linkType = self.gui.getLinkType()
        self.gliederung = self.gui.getGliederung()
        self.includeSub = self.gui.getIncludeSub()
        self.start = self.gui.getStart()
        self.end = self.gui.getEnd()

        pagenum = scribus.pageCount()
        for page in range(1, pagenum + 1):
            scribus.gotoPage(page)
            pageitems = scribus.getPageItems()
            for item in pageitems:
                if item[1] != 4:
                    continue
                self.textbox = item[0]
                self.evalTemplate()

        # hyphenate does not work
        # pagenum = scribus.pageCount()
        # for page in range(1, pagenum + 1):
        #     scribus.gotoPage(page)
        #     pageitems = scribus.getPageItems()
        #     for item in pageitems:
        #         if item[1] != 4:
        #             continue
        #         self.textbox = item[0]
        #         b = scribus.hyphenateText(self.textbox) # seems to have no effect!

        pagenum = scribus.pageCount()
        for page in range(1, pagenum + 1):
            scribus.gotoPage(page)
            pageitems = scribus.getPageItems()
            for item in pageitems:
                if item[1] != 4:
                    continue
                self.textbox = item[0]
                while scribus.textOverflows(self.textbox):
                    self.createNewPage()

        ausgabedatei = self.ausgabedatei
        if ausgabedatei == None or ausgabedatei == "":
            ausgabedatei = "ADFC_" + self.gliederung + (
                "_I_" if self.includeSub else "_") + self.start + "-" + self.end + "_" + self.linkType[0] + ".sla"
        try:
            scribus.saveDocAs(ausgabedatei)
        except Exception as e:
            print("Ausgabedatei", ausgabedatei, "konnte nicht geschrieben werden")
            raise e
        finally:
            self.touren = []
            self.termine = []
            self.gui.destroy()
            self.gui.quit()

    def evalTemplate(self):
        pos1 = 0
        while True:
            textlen = scribus.getTextLength(self.textbox)
            if textlen == 0:
                return
            scribus.selectText(0, textlen, self.textbox)
            alltext = unicode(scribus.getAllText(self.textbox))
            #logger.debug("alltext: %s %s", type(alltext), alltext)

            pos1 = alltext.find("/template", pos1)
            logger.debug("pos /template=%d", pos1)
            if pos1 < 0:
                return
            pos2 = alltext.find("/endtemplate")
            if pos2 < 0:
                raise Exception("kein /endtemplate nach /template")
            pos2 += 12   # len("/endtemplate")
            lines = alltext[pos1:pos2].split('\r')
            logger.debug("lines:%s %s", type(lines), str(lines))
            line0 = lines[0]
            lineN = lines[-1]
            #logger.debug("lineN: %s %s", type(lineN), lineN)
            if lineN != "/endtemplate":
                raise ValueError("Die letzte Zeile des templates darf nur /endtemplate enthalten")
            logger.debug("5et")
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

    def evalRun(self, tour):
        # logger.debug("evalRun1 %s", self.run.text)
        lines = self.run.text.split('\r')
        nl = ""
        for line in lines:
            if not line.startswith("/template") and\
                    not line.startswith("/endtemplate"):
                self.insertText(nl, self.run)
                self.expandLine(line, tour)
                nl = "\n"

    def expandLine(self, s, tour):
        #logger.debug("expand1 <%s>", s)
        spos = 0
        while True:
            mp = paramRE.search(s, spos)
            if mp is None:
                #logger.debug("noexp %s", s[spos:])
                self.insertText(s[spos:], self.run)
                return
            gp = mp.group(1).lower()
            #logger.debug("expand2 %s", gp)
            sp = mp.span()
            self.insertText(s[spos:sp[0]], self.run)
            mf = fmtRE.search(s, pos=spos)
            if mf is not None and sp[1] == mf.span()[0]: # i.e. if ${param] is followed immediately by .fmt()
                gf = mf.group(1)
                sf = mf.span()
                spos = sf[1]
                expanded = self.expandParam(gp, tour, gf)
            else:
                expanded = self.expandParam(gp, tour, None)
                spos = sp[1]
            #logger.debug("expand3 <%s>", str(expanded))
            if expanded is None: # special case for beschreibung, handled as markdown
                return
            if isinstance(expanded, list): # list is n runs + 1 string
                for run in expanded:
                    if isinstance(run, ScrbRun):
                        self.insertText(run.text, run)
                    else:
                        self.insertText(run, self.run)
            else:
                self.insertText(expanded, self.run)

    def expandParam(self, param, tour, format):
        try:
            f = self.expFunctions[param]
            return f(tour, format)
        except Exception as e:
            err = 'Fehler mit dem Parameter "' + param + \
                  '" des Events ' + self.tourMsg
            logger.exception(err)
            return param

    def expHeute(self, _, format):
        if format is None:
            return str(datetime.date.today())
        else:
            #return datetime.date.today().strftime(format)
            return datetime.datetime.now().strftime(format)

    def expStart(self, tour, format):
        dt = convertToMEZOrMSZ(tour.getDatumRaw())
        if format is None:
            return str(dt)
        else:
            return dt.strftime(format)

    def expEnd(self, tour, format):
        dt = convertToMEZOrMSZ(tour.getEndDatumRaw())
        if format is None:
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
        titel = tour.getTitel()
        logger.info("Titel: %s URL: %s", titel, self.url)
        return tour.getTitel()

    def expBeschreibung(self, tour, _):
        desc = tour.eventItem.get("description")
        desc = tourRest.removeSpcl(desc)
        desc = tourRest.removeHTML(desc)
        # did I ever need this?
        #desc = codecs.decode(desc, encoding = "unicode_escape")
        #logger.debug("desc type:%s <<<%s>>>", type(desc), desc)
        self.md.convert(desc)
        self.md.reset()
        return None

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
        self.run.cstyle = "WD2_Y_"
        return schwierigkeitMMap[tour.getSchwierigkeit()]

    def expTourLength(self, tour, _):
        return tour.getStrecke()

    def expPersonen(self, bezeichnung, tour):
        tl = tour.getPersonen()
        if len(tl) == 0:
            return ""
        run = copy.copy(self.run)
        run.cstyle = "ArialBold"
        run.text = bezeichnung + ": "
        return [ run, ", ".join(tl)]

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

        run = copy.copy(self.run)
        run.cstyle = "ArialBold"
        run.text = "Ort" + ("" if len(afs) == 1 else "e") + ": "
        return [ run, ", ".join(afl)]

    def expZusatzInfo(self, tour, _):
        zi = tour.getZusatzInfo()
        if len(zi) == 0:
            return None
        runs = []
        for z in zi:
            x = z.find(':') + 1
            run = copy.deepcopy(self.run)
            run.cstyle = "ArialBold"
            run.text = z[0:x] + " "
            runs.append(run)
            run = copy.deepcopy(self.run)
            run.text = z[x + 1:] + "\n"
            runs.append(run)
        return runs

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

def main():
    try:
        scribus.statusMessage('Running script...')
        #scribus.progressReset()
        adfc_gui.main("-scribus")
    finally:
        if scribus.haveDoc() > 0:
            scribus.redrawAll()
        scribus.statusMessage('Done.')
        #scribus.progressReset()

if __name__ == "__main__":
    import cProfile
    cProfile.run("main()", "cprof.prf")
    import pstats
    with open("cprof.txt", "w") as cprf:
        p = pstats.Stats("cprof.prf", stream=cprf)
        p.strip_dirs().sort_stats("cumulative").print_stats(20)
