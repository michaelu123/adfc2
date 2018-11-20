# encoding: utf-8

import tourRest
import os,sys
import json
import re
import datetime
import time
import markdown
from myLogger import logger
import docx

schwierigkeitMap = { 0: "sehr einfach", 1: "sehr einfach", 2: "einfach", 3: "mittel", 4: "schwer", 5: "sehr schwer"}
paramRE = re.compile(r"\${(\w*?)}")
fmtRE = re.compile(r"\.fmt\((.*?)\)")
strokeRE = r'(\~{2})(.+?)\1'
headerFontSizes = [ 0, 24, 18, 14, 12, 10, 8 ] # h1-h6 headers have fontsizes 24-8



"""
see https://stackoverflow.com/questions/4770297/convert-utc-datetime-string-to-local-datetime-with-python
"""
def convertToMEZOrMSZ(s: str):  # '2018-04-29T06:30:00+00:00'
    dt = time.strptime(s, "%Y-%m-%dT%H:%M:%S%z")
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

class DocxTreeHandler(markdown.treeprocessors.Treeprocessor):
    def __init__(self, md):
        super().__init__(md)
        self.ancestors = []
        self.states = []
        self.nodeHandler = { "h1": self.h1, "h2": self.h2, "h3": self.h3, "h4": self.h4, "h5": self.h5, "h6": self.h6,
            "p": self.p, "strong": self.strong, "em": self.em, "blockquote": self.blockQuote, "stroke": self.stroke,
            "ul": self.ul, "ol": self.ol, "li": self.li, "a": self.a, "hr": self.hr }

    def run(self, root):
        for child in root: # skip <div> root
            self.walkOuter(child)
        root.clear()

    def setDeps(self, docxHandler):
        self.docxHandler = docxHandler

    def printLines(self, s):
        while len(s) > 0:
            x = s.find('\n')
            if x >= 0:
                self.docxHandler.para.add_run(s[0:x])
                self.docxHandler.simpleNl()
                s = s[x + 1:]
            else:
                self.docxHandler.para.add_run(s)
                s = ""

    def walkOuter(self, node):
        global debug
        if debug:
            print(" " * self.lvl,"<<<<")
        try:
            self.nodeHandler[node.tag](node)
            if not node.tail is None:
                self.printLines(node.tail)
        except Exception as e:
            logger.exception("error in tour description")
        if debug:
            print(" " * self.lvl,">>>>")

    def walkInner(self, node):
            text = node.text
            tail = node.tail
            if text != None:
                ltext = text.replace("\n", "<nl>")
            else:
                ltext = "None"
            if tail != None:
                ltail = tail.replace("\n", "<nl>")
            else:
                ltail = "None"
            global debug
            if debug:
                print(" " * self.lvl, "node=", node.tag, ",text=", ltext, "tail=", ltail)
            if not text is None:
                self.printLines(text)
            for dnode in node:
                self.lvl += 4
                self.walkOuter(dnode)
                self.lvl -= 4

    def h1(self, node):
        sav = self.docxHandler.curStyle.size
        self.docxHandler.curStyle.size = headerFontSizes[1]
        self.walkInner(node)
        self.docxHandler.curStyle.size = sav
    def h2(self, node):
        sav = self.docxHandler.curStyle.size
        self.docxHandler.curStyle.size = headerFontSizes[2]
        self.walkInner(node)
        self.docxHandler.curStyle.size = sav
    def h3(self, node):
        sav = self.docxHandler.curStyle.size
        self.docxHandler.curStyle.size = headerFontSizes[3]
        self.walkInner(node)
        self.docxHandler.curStyle.size = sav
    def h4(self, node):
        sav = self.docxHandler.curStyle.size
        self.docxHandler.curStyle.size = headerFontSizes[4]
        self.walkInner(node)
        self.docxHandler.curStyle.size = sav
    def h5(self, node):
        sav = self.docxHandler.curStyle.size
        self.docxHandler.curStyle.size = headerFontSizes[5]
        self.walkInner(node)
        self.docxHandler.curStyle.size = sav
    def h6(self, node):
        sav = self.docxHandler.curStyle.size
        self.docxHandler.curStyle.size = headerFontSizes[6]
        self.walkInner(node)
        self.docxHandler.curStyle.size = sav
    def p(self, node):
        self.docxHandler.fontStyles = ""
        self.walkInner(node)
    def strong(self, node):
        sav = self.docxHandler.fontStyles
        self.docxHandler.fontStyles += "B"
        self.walkInner(node)
        self.docxHandler.fontStyles = sav
    def stroke(self, node):
        sav = self.docxHandler.fontStyles
        self.docxHandler.fontStyles += "U"
        self.walkInner(node)
        self.docxHandler.fontStyles = sav
    def em(self, node):
        sav = self.docxHandler.fontStyles
        self.docxHandler.fontStyles += "I"
        self.walkInner(node)
        self.docxHandler.fontStyles = sav
    def ul(self, node):
        self.numbered = False
        self.indent += 1
        self.walkInner(node)
        self.indent -= 1
    def ol(self, node):
        self.numbered = True
        sav = self.counter
        self.counter = 1
        self.indent += 1
        self.walkInner(node)
        self.indent -= 1
        self.counter = sav
    def li(self, node):
        if self.numbered:
            text = "  " * (self.indent * 3) + str(self.counter) + ". "
            self.counter += 1
        else:
            text = "  " * (self.indent * 3) + "\u25aa "
        sav = self.docxHandler.indentX
        self.docxHandler.indentX = 0.0
        self.printLines(text)
        self.docxHandler.indentX = 0
        self.walkInner(node)
        self.docxHandler.indentX = sav
    def a(self, node):
        self.docxHandler.url = node.attrib["href"]
        sav = self.docxHandler.curStyle.color
        self.docxHandler.curStyle.color = "238,126,13"
        self.walkInner(node)
        self.docxHandler.curStyle.color = sav
        self.docxHandler.url = None
    def blockQuote(self, node):
        node.text = node.tail = None
        sav = self.docxHandler.align
        if len(node) == 0: # multi_cell always does a newline et the end
            self.docxHandler.align = 'J'
        self.walkInner(node)
        self.docxHandler.align = sav
    def hr(self, node):
        self.docxHandler.extraNl()
        self.docxHandler.extraNl()

class DocxExtension(markdown.Extension):
    def extendMarkdown(self, md):
        self.docxTreeHandler = DocxTreeHandler(md)
        md.treeprocessors.register(self.docxTreeHandler, "docxtreehandler", 5)
        md.inlinePatterns.register(markdown.inlinepatterns.SimpleTagInlineProcessor(strokeRE, 'stroke'), 'stroke', 40)

def delete_paragraph(paragraph):
    # https://github.com/python-openxml/python-docx/issues/33
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

class DocxHandler:
    def __init__(self, gui):
        self.gui = gui
        self.terminselections = {}
        self.tourselections = {}
        self.touren = []
        self.termine = []
        self.url = None
        global debug
        try:
            lvl = os.environ["DEBUG"]
            debug = True
        except:
            debug = False
        self.docxExtension = DocxExtension()
        self.md = markdown.Markdown(extensions=[self.docxExtension])
        self.gui.docxTemplateName = r"C:\Users\Michael\PycharmProjects\ADFC1\venv\src\template.docx"
        if self.gui.docxTemplateName is None or self.gui.docxTemplateName == "":
            self.gui.docxTemplate()
        if self.gui.docxTemplateName is None or self.gui.docxTemplateName == "":
            raise ValueError("must specify path to .docx template!")
        self.doc = docx.Document(self.gui.docxTemplateName)
        self.docxExtension.docxTreeHandler.setDeps(self)
        self.parseParams()
        self.selFunctions = {
            "titelenthält": self.selTitelEnthält,
            "titelenthältnicht": self.selTitelEnthältNicht,
            "terminnr": self.selTourNr,
            "nichtterminnr": self.selNotTourNr,
            "tournr": self.selTourNr,
            "nichttournr": self.selNotTourNr,
            "radtyp": self.selRadTyp,
            "kategorie": self.selKategorie,
            "merkmal": self.selMerkmal,
        }
        self.expFunctions = {
            "heute": self.expHeute,
            "start": self.expStart,
            "end": self.expEnd,
            "nummer": self.expNummer,
            "titel": self.expTitel,
            "beschreibung": self.expBeschreibung,
            "abfahrten": self.expAbfahrten,
            "tourleiter": self.expTourLeiter,
            "betreuer": self.expBetreuer,
            "name": self.expName,
            "city": self.expCity,
            "street": self.expStreet,
            "kategorie": self.expKategorie,
            "schwierigkeit": self.expSchwierigkeit,
            "tourlänge": self.expTourLength,
            "abfahrten": self.expAbfahrten,
            "zusatzinfo": self.expZusatzInfo
        }

    def nothingFound(self):
        logger.info("Nichts gefunden")
        print("Nichts gefunden")

    def parseParams(self):
        texts = []
        for para in self.doc.paragraphs:
            delete_paragraph(para)
            if para.text.find("-----") >= 0:
                break;
            if para.text == "" or para.text.startswith("Kommentar"):
                continue
            if para.style.name.startswith("List"):
                continue
            if para.paragraph_format.left_indent != None:
                continue
            texts.append(para.text)
        lines = "\n".join(texts).split('\n')
        texts = None
        #defaults:
        self.linkType = "frontend"
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
                    self.linkType = words[1].lower()
                    lx += 1
                elif word0 == "ausgabedatei":
                    self.ausgabedatei = words[1]
                    lx += 1
                else:
                    raise ValueError(
                        "Unbekannter Parameter " + word0 + ", erwarte linktyp oder ausgabedatei")
            elif word0 not in ["selektion", "terminselektion", "tourselektion" ]:
                raise ValueError(
                    "Unbekannter Parameter " + word0 + ", erwarte selektion, terminselektion oder tourselektion")
            else:
                lx = self.parseSel(word0, lines, lx+1, selections)
        selection = selections.get("selektion")
        self.gliederung = selection.get("gliederungen")
        if self.gliederung is None or self.gliederung == "":
            self.gliederung = self.gui.getGliederung()
        self.includeSub = selection.get("mituntergliederungen") == "ja"
        if self.includeSub is None:
            self.includeSub = self.gui.getIncludeSub()
        self.start = selection.get("beginn")
        if self.start is None or self.start == "":
            self.start = self.gui.getStart()
        self.end = selection.get("ende")
        if self.end is None or self.end == "":
            self.end = self.gui.getEnd()
        sels = selections.get("terminselektion")
        for sel in sels.values():
            self.terminselections[sel.get("name")] = sel
            for key in sel.keys():
                if key != "name" and not isinstance(sel[key], list):
                    sel[key] = [ sel[key] ]
        sels = selections.get("tourselektion")
        for sel in sels.values():
            self.tourselections[sel.get("name")] = sel
            for key in sel.keys():
                if key != "name" and not isinstance(sel[key], list):
                    sel[key] = [ sel[key] ]

    def parseSel(self, word, lines, lx, selections):
        selections[word] = sel = sel2 = {}
        while lx < len(lines):
            line = lines[lx]
            if not line[0].isspace():
                return lx
            words = line.split()
            word0 = words[0].lower().replace(":", "")
            if word0 == "name":
                word1 = words[1].lower()
                sel[word1] = sel2 = {}
                sel2["name"] = word1
            else:
                if len(words) == 2:
                    sel2[word0] = words[1]
                else:
                    sel2[word0] = "".join(words[1:]).split(",")
            lx += 1
        return lx

    def parseDoc(self):
        texts = []
        for para in self.doc.paragraphs:
            print("<<< " + para.text + " >>>")

    def getIncludeSub(self):
        return self.includeSub
    def getType(self):
        if len(self.terminselections) != 0 and len(self.tourselections) != 0:
            return "Alles";
        if len(self.terminselections) != 0:
            return "Termin"
        if len(self.tourselections) != 0:
            return "Radtour"
        return self.gui.getTyp()
    def getRadTyp(self):
        rts = set()
        for sel in self.tourselections.values():
            l = sel.get("radtyp")
            if l is None or len(l) == 0:
                l = [ self.gui.getRadTyp() ]
            for elem in l:
                rts.add(elem)
        if "Alles" in rts:
            return "Alles";
        if len(rts) == 1:
            return rts[0]
        return "Alles"
    def getUnitKeys(self):
        return self.gliederung
    def getStart(self):
        return self.start
    def getEnd(self):
        return self.end

    def handleTour(self, tour):
        self.touren.append(tour)
    def handleTermin(self, tour):
        self.termine.append(tour)
    def handleEnd(self):
        print("Template", self.gui.docxTemplateName, "wird abgearbeitet")
        if self.linkType == None or self.linkType == "":
            self.linkType = self.gui.getLinkType()
        self.parseDoc()

        """
        lines = self.pdfJS.get("text");
        lineCnt = len(lines)
        lineNo = 0
        self.pdf.set_x(self.margins[0]) # left
        self.pdf.set_y(self.margins[1]) # top
        self.setStyle(self.styles.get("body"))
        self.pdf.cell(w=0, h=10, txt="", ln=1)
        while lineNo < lineCnt:
            line = lines[lineNo]
            if line.startswith("/comment"):
                lineNo += 1
                continue
            if line.startswith("/template"):
                t1 = lineNo
                lineNo += 1
                while not lines[lineNo].startswith("/endtemplate"):
                    lineNo += 1
                t2 = lineNo
                lineNo += 1
                tempLines = lines[t1:t2] # /endtemplate not included
                self.evalTemplate(tempLines)
            else:
                self.evalLine(line, None)
                lineNo += 1
        """
        if self.ausgabedatei == None or self.ausgabedatei == "":
            self.ausgabedatei = self.gui.docxTemplateName.rsplit(".", 1)[0] + "_" + self.linkType[0] + ".docx"
        self.doc.save(self.ausgabedatei)
        print("Ausgabedatei", self.ausgabedatei, "wurde erzeugt")
        try:
            opath = os.path.abspath(self.ausgabedatei)
            os.startfile(opath)
        except Exception as e:
            logger.exception("opening " + self.ausgabedatei)

    def simpleNl(self):
        pass
    def extraNl(self):
        self.simpleNl()

    def evalLine(self, line, tour):
        if line.strip() == "":
            self.extraNl()
            return
        global debug
        if debug:
            print("line", line)
        text = []
        self.align = "L"
        self.fontStyles = ""
        self.curStyle = self.styles.get("body")
        self.indentX = 0.0
        words = line.split()
        l = len(words)
        last = l - 1
        for i in range(l):
            word = words[i];
            if word.startswith("/"):
                cmd = word[1:]
                if cmd in self.styles.keys():
                    self.handleText("".join(text), tour)
                    text = []
                    self.curStyle = self.styles.get(cmd)
                elif cmd in ["right", "left", "center", "block"]:
                    self.handleText("".join(text), tour)
                    text = []
                    self.align = cmd[0].upper()
                    if self.align == 'B':
                        self.align = 'J' ## justification
                elif cmd in ["bold", "italic", "underline"]:
                    self.handleText("".join(text), tour)
                    text = []
                    self.fontStyles += cmd[0].upper()
                else:
                    if i < last:
                        word = word + " "
                    text.append(word)
            else:
                word = word.replace("\uaffe", "\n")
                if i < last:
                    word = word + " "
                text.append(word)
        self.handleText("".join(text), tour)
        self.simpleNl()

    def evalTemplate(self, lines):
        global debug
        if debug:
            print("template:")
        words = lines[0].split()
        typ = words[1]
        if typ != "/tour" and typ != "/termin":
            raise ValueError("second word after /template must be /tour or /termin")
        typ = typ[1:]
        sel = words[2]
        if not sel.startswith("/selection="):
            raise ValueError("third word after /template must start with /selection=")
        sel = sel[11:]
        sels = self.tourselections if typ == "tour" else self.terminselections
        if not sel in sels:
            raise ValueError("selection " + sel + " not in " + typ + "selections")
        sel = sels[sel]
        touren = self.touren if typ == "tour" else self.termine
        self.evalTouren(sel, touren, lines[1:])

    def evalTouren(self, sel, touren, lines):
        selectedTouren = []
        for tour in touren:
            if self.selected(tour, sel):
                selectedTouren.append(tour)
        if len(selectedTouren) == 0:
            return
        lastTour = selectedTouren[-1]
        for tour in selectedTouren:
            for line in lines:
                if line.startswith("/comment"):
                    continue
                self.evalLine(line, tour)
            if tour != lastTour: # extra line between touren, not after the last one
                self.evalLine("", None)


    def selTitelEnthält(self, tour, lst):
        titel = tour.getTitel()
        for  elem in lst:
            if titel.find(elem) > 0:
                return True
        return False

    def selTitelEnthältNicht(self, tour, lst):
        titel = tour.getTitel()
        for  elem in lst:
            if titel.find(elem) > 0:
                return False
        return True

    def selRadTyp(self, tour, lst):
        if "Alles" in lst:
            return True
        radTyp = tour.getRadTyp()
        return radTyp in lst

    def selTourNr(self, tour, lst):
        nr = int(tour.getNummer())
        return nr in lst

    def selNotTourNr(self, tour, lst):
        nr = int(tour.getNummer())
        return not nr in lst

    def selKategorie(self, tour, lst):
        kat = tour.getKategorie()
        return kat in lst

    def selMerkmal(self, tour, lst):
        mm = tour.getMerkmal()
        return mm in lst

    def selected(self, tour, sel):
        for key in sel.keys():
            if key == "name" or key.startswith("comment"):
                continue
            try:
                f = self.selFunctions[key]
                lst = sel[key]
                if not f(tour, lst):
                    return False
            except Exception as e:
                logger.exception("no function for selection verb " + key + " in selection " + sel.get("name"))
        else:
            return True

    def expand(self, s, tour):
        while True:
            mp = paramRE.search(s)
            if mp == None:
                return s
            gp = mp.group(1)
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
                return ""
            try:
                s = s[0:sp[0]] + expanded + s[sp[1]:]
            except Exception as e:
                logger.error("expanded = " + expanded)
        return s

    def expandParam(self, param, tour, format):
        try:
            f = self.expFunctions[param]
            return f(tour, format)
        except Exception as e:
            logger.exception('error with parameter "' + param + '" for tour ' + tour.getTitel())
            return param

    def expHeute(self, tour, format):
        if format == None:
            return str(datetime.date.today())
        else:
            return datetime.date.today().strftime(format)

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

    def expNummer(self, tour, format):
        return tour.getRadTyp()[0].upper() + tour.getNummer()

    def expTitel(self, tour, format):
        if self.linkType == "frontend":
            self.url = tour.getFrontendLink()
        elif self.linkType == "backend":
            self.url = tour.getBackendLink()
        else:
            self.url = None
        logger.info("Titel: " + tour.getTitel())
        return tour.getTitel()

    def expBeschreibung(self, tour, format):
        desc = tour.eventItem.get("description")
        desc = tourRest.removeHTML(desc)
        #desc = codecs.decode(desc, encoding = "unicode_escape")
        self.md.convert(desc)
        self.md.reset()
        return None

    def expName(self, tour, format):
        return tour.getName()
    def expCity(self, tour, format):
        return tour.getCity()
    def expStreet(self, tour, format):
        return tour.getStreet()

    def expKategorie(self, tour, format):
        return tour.getKategorie()

    def expSchwierigkeit(self, tour, format):
        return schwierigkeitMap[tour.getSchwierigkeit()]

    def expTourLength(self, tour, format):
        return tour.getStrecke()

    def expTourLeiter(self, tour, format):
        tl = tour.getPersonen()
        if len(tl) == 0:
            return
        self.evalLine("/bold Tourleiter: /block " + "\uaffe".join(tl), tour)

    def expAbfahrten(self, tour, format):
        afs = tour.getAbfahrten()
        if len(afs) == 0:
            return
        afl = [ af[0] + " " + af[1] + " " + af[2] for af in afs]
        self.evalLine("/bold Ort" + ("" if len(afs) == 1 else "e") + ": /block " + "\uaffe".join(afl), tour)

    def expBetreuer(self, tour, format):
        tl = tour.getPersonen()
        if len(tl) == 0:
            return
        self.evalLine("/bold Betreuer: /block " + "\uaffe".join(tl), tour)

    def expZusatzInfo(self, tour, format):
        zi = tour.getZusatzInfo()
        if len(zi) == 0:
            return
        self.evalLine("/bold Zusatzinfo: /block " + "\uaffe".join(zi), tour)

