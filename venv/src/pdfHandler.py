# encoding: utf-8

import tourRest
import os,sys
import json
import re
import datetime
import time
import markdown
from myLogger import logger
from fpdf import FPDF

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

class Style:
    def __init__(self, name:str, type:str, font:str, fontStyle:str, size:int, color:str, dimen:str):
        self.name = name
        self.type = type
        self.font = font
        self.fontStyle = fontStyle
        self.size = size
        self.color = color
        self.dimen = dimen
    def copy(self):
        return Style(self.name, self.type, self.font, self.fontStyle, self.size, self.color, self.dimen)
    def __str__(self):
        return self.name

class PDFTreeHandler(markdown.treeprocessors.Treeprocessor):
    def __init__(self, md):
        super().__init__(md)
        self.ancestors = []
        self.states = []
        self.nodeHandler = { "h1": self.h1, "h2": self.h2, "h3": self.h3, "h4": self.h4, "h5": self.h5, "h6": self.h6,
            "p": self.p, "strong": self.strong, "em": self.em, "blockquote": self.blockQuote, "stroke": self.stroke,
            "ul": self.ul, "ol": self.ol, "li": self.li, "a": self.a, "hr": self.hr }

    def run(self, root):
        self.pdfHandler.setStyle(self.pdfHandler.styles.get("body").copy()) # now == curStyle
        self.pdfHandler.align = "L"
        self.indent = 0
        self.pdfHandler.indentX = 0.0
        self.lvl = 4
        self.counter = 0
        for child in root: # skip <div> root
            self.walkOuter(child)
        root.clear()

    def setDeps(self, pdfHandler):
        self.pdfHandler = pdfHandler
        self.pdf = pdfHandler.pdf

    def printLines(self, s):
        while len(s) > 0:
            x = s.find('\n')
            if x >= 0:
                self.pdfHandler.handleText(s[0:x], None)
                self.pdfHandler.simpleNl()
                s = s[x + 1:]
            else:
                self.pdfHandler.handleText(s, None)
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
        sav = self.pdfHandler.curStyle.size
        self.pdfHandler.curStyle.size = headerFontSizes[1]
        self.walkInner(node)
        self.pdfHandler.curStyle.size = sav
    def h2(self, node):
        sav = self.pdfHandler.curStyle.size
        self.pdfHandler.curStyle.size = headerFontSizes[2]
        self.walkInner(node)
        self.pdfHandler.curStyle.size = sav
    def h3(self, node):
        sav = self.pdfHandler.curStyle.size
        self.pdfHandler.curStyle.size = headerFontSizes[3]
        self.walkInner(node)
        self.pdfHandler.curStyle.size = sav
    def h4(self, node):
        sav = self.pdfHandler.curStyle.size
        self.pdfHandler.curStyle.size = headerFontSizes[4]
        self.walkInner(node)
        self.pdfHandler.curStyle.size = sav
    def h5(self, node):
        sav = self.pdfHandler.curStyle.size
        self.pdfHandler.curStyle.size = headerFontSizes[5]
        self.walkInner(node)
        self.pdfHandler.curStyle.size = sav
    def h6(self, node):
        sav = self.pdfHandler.curStyle.size
        self.pdfHandler.curStyle.size = headerFontSizes[6]
        self.walkInner(node)
        self.pdfHandler.curStyle.size = sav
    def p(self, node):
        self.pdfHandler.fontStyles = ""
        self.walkInner(node)
    def strong(self, node):
        sav = self.pdfHandler.fontStyles
        self.pdfHandler.fontStyles += "B"
        self.walkInner(node)
        self.pdfHandler.fontStyles = sav
    def stroke(self, node):
        # see below changed code of pyfpdf
        sav = self.pdfHandler.fontStyles
        self.pdfHandler.fontStyles += "U"
        self.walkInner(node)
        self.pdfHandler.fontStyles = sav
    def em(self, node):
        sav = self.pdfHandler.fontStyles
        self.pdfHandler.fontStyles += "I"
        self.walkInner(node)
        self.pdfHandler.fontStyles = sav
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
        sav = self.pdfHandler.indentX
        self.pdfHandler.indentX = 0.0
        self.printLines(text)
        self.pdfHandler.indentX = self.pdf.get_x()
        self.walkInner(node)
        self.pdfHandler.indentX = sav
    def a(self, node):
        self.pdfHandler.url = node.attrib["href"]
        sav = self.pdfHandler.curStyle.color
        self.pdfHandler.curStyle.color = "238,126,13"
        self.walkInner(node)
        self.pdfHandler.curStyle.color = sav
        self.pdfHandler.url = None
    def blockQuote(self, node):
        node.text = node.tail = None
        sav = self.pdfHandler.align
        if len(node) == 0: # multi_cell always does a newline et the end
            self.pdfHandler.align = 'J'
        self.walkInner(node)
        self.pdfHandler.align = sav
    def hr(self, node):
        self.pdfHandler.extraNl()
        x = self.pdf.get_x()
        y = self.pdf.get_y()
        self.pdf.line(x, y, self.pdfHandler.pageWidth - self.pdfHandler.margins[2], y)
        self.pdfHandler.extraNl()

class PDFExtension(markdown.Extension):
    def extendMarkdown(self, md):
        self.pdfTreeHandler = PDFTreeHandler(md)
        md.treeprocessors.register(self.pdfTreeHandler, "pdftreehandler", 5)
        md.inlinePatterns.register(markdown.inlinepatterns.SimpleTagInlineProcessor(strokeRE, 'stroke'), 'stroke', 40)

class PDFHandler:
    def __init__(self, gui):
        self.gui = gui
        self.styles = {}
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
        self.pdfExtension = PDFExtension()
        self.md = markdown.Markdown(extensions=[self.pdfExtension])
        if self.gui.pdfTemplateName is None or self.gui.pdfTemplateName == "":
            self.gui.pdfTemplate()
        if self.gui.pdfTemplateName is None or self.gui.pdfTemplateName == "":
            raise ValueError("must specify path to PDF template!")
        try:
            with open(self.gui.pdfTemplateName, "r", encoding="utf-8-sig") as jsonFile:
                self.pdfJS = json.load(jsonFile)
        except Exception as e:
            print("Wenn Sie einen decoding-Fehler bekommen, öffnen Sie " + self.gui.pdfTemplateName + " mit notepad, dann 'Speichern unter' mit Codierung UTF-8")
            raise e
        self.parseTemplate()
        self.selFunctions = {
            "titelenthält": self.selTitelEnthält,
            "titelenthältnicht": self.selTitelEnthältNicht,
            "radtyp": self.selRadTyp,
            "tournr": self.selTourNr,
            "kategorie": self.selKategorie,
            "nichttournr": self.selNotTourNr
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

    def parseTemplate(self):
        for key in ["pagesettings", "fonts", "styles", "selection", "text"]:
            if self.pdfJS.get(key) == None:
                raise ValueError("pdf template " + self.gui.pdfTemplateName + " must have a section " + key)
        pagesettings = self.pdfJS.get("pagesettings")
        leftMargin = pagesettings.get("leftmargin")
        rightMargin = pagesettings.get("rightmargin")
        topMargin = pagesettings.get("topmargin")
        bottomMargin = pagesettings.get("bottommargin")
        self.margins = (leftMargin, topMargin, rightMargin)
        self.linespacing = pagesettings.get("linespacing") # float
        self.linkType = pagesettings.get("linktype")
        self.ausgabedatei = pagesettings.get("ausgabedatei")
        orientation = pagesettings.get("orientation")[0].upper() # P or L
        orientation = pagesettings.get("orientation")[0].upper() # P or L
        format = pagesettings.get("format")
        self.pdf = FPDF(orientation, "mm", format)
        self.pageWidth = FPDF.get_page_format(format, 1.0)[0] / (72.0/25.4)
        self.pdf.add_page()
        self.pdf.set_margins(left=leftMargin, top=topMargin, right=rightMargin)
        self.pdf.set_auto_page_break(True, margin=bottomMargin)
        self.pdfExtension.pdfTreeHandler.setDeps(self)

        self.pdf.add_font("arialuc", "", pyinst("_builtin_fonts/arial.ttf"), True)
        self.pdf.add_font("arialuc", "B", pyinst("_builtin_fonts/arialbd.ttf"), True)
        self.pdf.add_font("arialuc", "BI", pyinst("_builtin_fonts/arialbi.ttf"), True)
        self.pdf.add_font("arialuc", "I", pyinst("_builtin_fonts/ariali.ttf"), True)

        fonts = self.pdfJS.get("fonts")
        for font in iter(fonts):
            family = font.get("family")
            if family is None or family == "":
                raise ValueError("font family not specified")
            file = font.get("file")
            if file is None or file == "":
                raise ValueError("font file not specified")
            fontStyle = font.get("fontstyle")
            if fontStyle is None:
                fontStyle = ""
            unicode = font.get("unicode")
            if unicode is None:
                unicode = True
            self.pdf.add_font(family, fontStyle, file, unicode)
        styles = self.pdfJS.get("styles")
        for style in iter(styles):
            name = style.get("name")
            if name is None or name == "":
                raise ValueError("style name not specified")
            type = style.get("type")
            if type is None:
                type = "text"
            font = style.get("font")
            if font is None and type != "image":
                raise ValueError("style font not specified")
            fontstyle = style.get("style")
            if fontstyle is None:
                fontstyle = ""
            size = style.get("size")
            if size is None and type != "image":
                raise ValueError("style size not specified")
            color = style.get("color")
            if color is None:
                color = "0,0,0" # black
            dimen = style.get("dimen")
            self.styles[name] = Style(name, type, font, fontstyle, size, color, dimen)
        selection = self.pdfJS.get("selection")
        self.gliederung = selection.get("gliederung")
        if self.gliederung is None or self.gliederung == "":
            self.gliederung = self.gui.getGliederung()
        self.includeSub = selection.get("includesub")
        if self.includeSub is None:
            self.includeSub = self.gui.getIncludeSub()
        self.start = selection.get("start")
        if self.start is None or self.start == "":
            self.start = self.gui.getStart()
        self.end = selection.get("end")
        if self.end is None or self.end == "":
            self.end = self.gui.getEnd()
        sels = selection.get("terminselection")
        for sel in iter(sels):
            self.terminselections[sel.get("name")] = sel
            for key in sel.keys():
                if key != "name" and not isinstance(sel[key], list):
                    sel[key] = [ sel[key] ]
        sels = selection.get("tourselection")
        for sel in iter(sels):
            self.tourselections[sel.get("name")] = sel
            for key in sel.keys():
                if key != "name" and not isinstance(sel[key], list):
                    sel[key] = [ sel[key] ]

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
        print("Template", self.gui.pdfTemplateName, "wird abgearbeitet")
        if self.linkType == None or self.linkType == "":
            self.linkType = self.gui.getLinkType()
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
        if self.ausgabedatei == None or self.ausgabedatei == "":
            self.ausgabedatei = self.gui.pdfTemplateName.rsplit(".", 1)[0] + "_" + self.linkType[0] + ".pdf"
        self.pdf.output(dest='F', name= self.ausgabedatei)
        print("Ausgabedatei", self.ausgabedatei, "wurde erzeugt")
        try:
            opath = os.path.abspath(self.ausgabedatei)
            os.startfile(opath)
        except Exception as e:
            logger.exception("opening " + self.ausgabedatei)

    def simpleNl(self):
        x = self.pdf.get_x()
        if x > self.margins[0]:
            self.pdf.ln()

    def extraNl(self):
        self.simpleNl()
        self.pdf.ln()

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

    def handleText(self, s:str, tour):
        s = self.expand(s, tour)
        if s == "":
            return
        #print("Text:", s)
        if self.curStyle.type == "image":
            self.drawImage(pyinst(s))
            return
        style = self.curStyle.copy()
        for fs in self.fontStyles:
            if style.fontStyle.find(fs) == -1:
                style.fontStyle += fs
        #self.fontStyles = ""
        self.setStyle(style)
        h=(style.size * 0.35278 + self.linespacing)
        if self.align == 'J':
            self.pdf.multi_cell(w=0, h=h, txt=s, border=0, align=self.align, fill=0)
        elif self.align == 'R':
            self.pdf.cell(w=0, h=h, txt=s, border=0, ln=0, align=self.align, fill=0, link=self.url)
        else:
            try:
                w = self.pdf.get_string_width(s)
            except Exception as e:
                pass
            x = self.pdf.get_x()
            if (x + w) >= (self.pageWidth - self.margins[2]): # i.e. exceeds right margin
                self.multiline(h, s)
            else:
                self.pdf.cell(w=w, h=h, txt=s, border=0, ln=0, align=self.align, fill=0, link=self.url)
            x = self.pdf.get_x()
        self.url = None

    def multiline(self, h: float, s: str):
        """ line too long, see if I can split line after blank """
        x = self.pdf.get_x()
        l = len(s)
        # TODO limit l so that we do not    search too long for a near enough blank
        while l > 0:
            w = self.pdf.get_string_width(s)
            if (x + w) < (self.pageWidth - 1 - self.margins[2]):
                self.pdf.cell(w=w, h=h, txt=s, border=0, ln=0, align=self.align, fill=0, link=self.url)
                x = self.pdf.get_x()
                return
            nlx = s.find("\n", 0, l)
            lb = s.rfind(' ', 0, l)  # last blank
            if nlx >= 0 and nlx < lb:
                lb = nlx + 1
            if lb == -1:  # can not split line
                if x > self.margins[0]:
                    self.pdf.ln()
                    if self.indentX > 0.0:
                        self.pdf.set_x(self.indentX)
                    x = self.pdf.get_x()
                    l = len(s)
                    continue
                else:  # emergency, can not split line
                    w = self.pdf.get_string_width(s)
                    self.pdf.cell(w=w, h=h, txt=s, border=0, ln=1, align=self.align, fill=0, link=self.url)
                    return
            sub = s[0:lb]
            w = self.pdf.get_string_width(sub)
            if (x + w) >= (self.pageWidth - 1 - self.margins[2]):
                l = lb
                continue
            self.pdf.cell(w=w, h=h, txt=sub, border=0, ln=0, align=self.align, fill=0, link=self.url)
            x = self.pdf.get_x()
            s = s[lb + 1:]
            w = self.pdf.get_string_width(s)
            if x > self.margins[0] and (x + w) >= (self.pageWidth - 1 - self.margins[2]):
                self.pdf.ln()
                if self.indentX > 0.0:
                    self.pdf.set_x(self.indentX)
                x = self.pdf.get_x()
            l = len(s)

    def setStyle(self, style:Style):
        #print("Style:", style)
        self.pdf.set_font(style.font, style.fontStyle, style.size)
        rgb = style.color.split(',')
        self.pdf.set_text_color(int(rgb[0]),int(rgb[1]),int(rgb[2]))

    def drawImage(self, imgName:str):
        style = self.curStyle
        dimen = style.dimen # 60x40, wxh
        wh = dimen.split('x')
        w = int(wh[0])
        h = int(wh[1])
        x = self.pdf.get_x()
        y = self.pdf.get_y()
        if self.align == 'R':
            x = self.pageWidth - self.margins[2] - w - 10
        y -= h  # align lower edge of image with baseline of text (or so)
        if y < self.margins[1]:
            y = self.margins[1]
        self.pdf.image(imgName.strip(), x=x, y=y, w=w) # h=h
        self.pdf.set_y(self.pdf.get_y() + 7)

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

    """
    changed code in fpdf/fpdf.py, to change underline to "stroke through letter" (i.e. line half a fontsize higher and thinner)
    def _dounderline(self, x, y, txt):
        #Underline text
        up=self.current_font['up']
        ut=self.current_font['ut']
        w=self.get_string_width(txt, True)+self.ws*txt.count(' ')
        y -= self.font_size / 2.0 # MUH this line added to stroke through letter instead underline
        return sprintf('%.2f %.2f %.2f %.2f re f',x*self.k,(self.h-(y-up/1000.0*self.font_size))*self.k,w*self.k,-ut/2000.0*self.font_size_pt)
        # MUH 2000 was 1000, seems to be line thickness
    """

