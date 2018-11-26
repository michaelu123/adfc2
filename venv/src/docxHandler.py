# encoding: utf-8

import tourRest
import selektion
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

def delete_paragraph(paragraph):
    # https://github.com/python-openxml/python-docx/issues/33
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

def delete_run(run):
    r = run._element
    r.getparent().remove(r)
    r._p = r._element = None

def add_run_copy(paragraph, run, text=None):
    newr = paragraph.add_run(text=run.text if text is None else text, style=run.style)
    newr.bold = run.bold
    newr.italic = run.italic
    newr.underline = run.underline
    newr.font.all_caps = run.font.all_caps
    newr.font.bold = run.font.bold
    if run.font.color != None and run.font.color.rgb != None:
        newr.font.color.rgb = run.font.color.rgb
    if run.font.color != None and run.font.color.theme_color != None:
        newr.font.color.theme_color = run.font.color.theme_color
    #newr.font.color.type = run.font.color.type is readonly
    newr.font.complex_script = run.font.complex_script
    newr.font.cs_bold = run.font.cs_bold
    newr.font.cs_italic = run.font.cs_italic
    newr.font.double_strike = run.font.double_strike
    newr.font.emboss = run.font.emboss
    newr.font.hidden = run.font.hidden
    newr.font.highlight_color = run.font.highlight_color
    newr.font.imprint = run.font.imprint
    newr.font.italic = run.font.italic
    newr.font.math = run.font.math
    newr.font.name = run.font.name
    newr.font.no_proof = run.font.no_proof
    newr.font.outline = run.font.outline
    newr.font.rtl = run.font.rtl
    newr.font.shadow = run.font.shadow
    newr.font.size = run.font.size
    newr.font.small_caps = run.font.small_caps
    newr.font.snap_to_grid = run.font.snap_to_grid
    newr.font.spec_vanish = run.font.spec_vanish
    newr.font.strike = run.font.strike
    newr.font.subscript = run.font.subscript
    newr.font.superscript = run.font.superscript
    newr.font.underline = run.font.underline
    newr.font.web_hidden = run.font.web_hidden
    return newr

def insert_paragraph_copy_before(paraBefore, para):
    newp = paraBefore.insert_paragraph_before()
    newp.alignment = para.alignment
    newp.style = para.style
    newp.paragraph_format.alignment = para.paragraph_format.alignment
    newp.paragraph_format.first_line_indent = para.paragraph_format.first_line_indent
    newp.paragraph_format.keep_together = para.paragraph_format.keep_together
    newp.paragraph_format.keep_with_next = para.paragraph_format.keep_with_next
    newp.paragraph_format.left_indent = para.paragraph_format.left_indent
    newp.paragraph_format.line_spacing = para.paragraph_format.line_spacing
    newp.paragraph_format.line_spacing_rule = para.paragraph_format.line_spacing_rule
    newp.paragraph_format.page_break_before = para.paragraph_format.page_break_before
    newp.paragraph_format.right_indent = para.paragraph_format.right_indent
    newp.paragraph_format.space_after = para.paragraph_format.space_after
    newp.paragraph_format.space_before = para.paragraph_format.space_before
    for ts in para.paragraph_format.tab_stops:
        newp.paragraph_format.add_tab_stop(ts-position, ts.alignment, ts.leader)
    newp.paragraph_format.widow_control = para.paragraph_format.widow_control
    for run in para.runs:
        add_run_copy(newp, run)
    return newp

def eqFont(f1, f2):
    if f1.name != f2.name:
        return False
    if f1.size != f2.size:
        return False
    return True

def eqStyle(s1, s2):
    if s1.name != s2.name:
        return False
    return True

def eqColor(r1, r2):
    p1 = hasattr(r1._element, "rPr")
    p2 = hasattr(r2._element, "rPr")
    if not p1 and not not p2:
        return True
    if p1 and not p2:
        return False
    if not p1 and p2:
        return False
    p1 = hasattr(r1._element.rPr, "color")
    p2 = hasattr(r2._element.rPr, "color")
    if not p1 and not p2:
        return True
    if p1 and not p2:
        return False
    if not p1 and p2:
        return False
    try:
        c1 = r1._element.rPr.color
        c2 = r2._element.rPr.color
    except:
        print("!!")
    if c1 == None and c2 == None:
        return True
    if c1 != None and c2 == None:
        return False
    if c1 == None and c2 != None:
        return False
    return c1.val == c2.val

def split_run(para, runs, run, x):
    runX = runs.index(run) + 1
    t1 = run.text[0:x]
    t2 = run.text[x:]
    run.text = t1
    new_run = add_run_copy(para, run, text=t2)
    # the insert does not work as expected, the new_run is always inserted into the same place,
    # irrespective of runX
    #    para._p.insert(runX, new_run._r)
    # therefore, we append all runs behind runX AFTER the newly appended run
    # i.e. we copy a b t1 c d t2 to a b t1 t2 c d, by appending c, d
    # this is all trial and error, and completely obscure...
    while runX < len(runs):
        para._p.append(runs[runX]._r)
        runX += 1
    print("splitRes:", " ".join(["<" + run.text + ">" for run in para.runs]))

"""
    This function combines the texts of successive runs with same font,style,color into one run.
    Word splits for unknown reasons continuous text like "Kommentar" into two runs "K" and "ommentar"!?
    Our parameters ${name} are split int "${", "name", "}". This makes parsing too difficult, so we combine first.
    But then we may have several ${param}s within one run. We then split the runs again so that each parameter is 
    in its own run.
"""
def combineRuns(doc):
    paras = doc.paragraphs
    for para in paras:
        print("Para ", str(para), para.text, " align:", para.alignment, "style:", para.style.name)
        runs = para.runs
        prevRun = None
        for run in runs:
            print("Run '", run.text, "' bold:", run.bold, " font:", run.font.name, run.font.size, " style:", run.style.name)
            if prevRun != None and prevRun.bold == run.bold and prevRun.italic == run.italic and \
                prevRun.underline == run.underline and \
                eqColor(prevRun, run) and \
                eqFont(prevRun.font, run.font) and \
                eqStyle(prevRun.style, run.style):
                prevRun.text += run.text
                delete_run(run)
            else:
                prevRun = run
    paras = doc.paragraphs
    for para in paras:
        if para.text.find("${") > 0:
            splitted = True
            while splitted:
                splitted = False
                runs = para.runs
                for run in runs:
                    mp = paramRE.search(run.text, 1)
                    if mp == None:
                        continue
                    sp = mp.span()
                    split_run(para, runs, run, sp[0])
                    splitted = True
                    break

def add_hyperlink_into_run(paragraph, run, i, url):
    runs = paragraph.runs
    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )
    hyperlink.append(run._r)
    paragraph._p.insert(i+1, hyperlink)

def move_run_before(i, paragraph, run):
    paragraph._p.insert(i, run._r)
    print()

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
        combineRuns(self.doc)
        self.docxExtension.docxTreeHandler.setDeps(self)
        self.parseParams()
        self.selFunctions = selektion.getSelFunctions()
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
                lst = ",".join(words[1:]).split(",")
                sel2[word0] = lst[0] if len(lst) == 1 else lst
            lx += 1
        return lx

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
        paragraphs = self.doc.paragraphs
        paraCnt = len(paragraphs)
        paraNo = 0
        while paraNo < paraCnt:
            para = paragraphs[paraNo]
            if para.text.startswith("/template"):
                p1 = paraNo
                while True:
                    if para.text.find("/endtemplate") >= 0:
                        break
                    delete_paragraph(para)
                    paraNo += 1
                    para = paragraphs[paraNo]
                delete_paragraph(para)
                p2 = paraNo
                paraNo += 1
                self.paraBefore = None if paraNo == paraCnt else paragraphs[paraNo]
                tempParas = paragraphs[p1:p2+1]
                self.evalTemplate(tempParas)
            else:
                self.evalPara(para)
                paraNo += 1

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

    def evalPara(self, para):
        global debug
        if debug:
            print("para", para.text)
        for run in para.runs:
            self.evalRun(run, None)
        self.simpleNl()

    def evalTemplate(self, paras):
        global debug
        if debug:
            print("template:")
        para0 = paras[0]
        para0Lines = para0.runs[0].text.split('\n')
        if not para0Lines[0].startswith("/template"):
            raise ValueError("/template muss am Anfang der ersten Zeile eines Paragraphen stehen")
        paraN = paras[-1]
        paraNLines = paraN.runs[-1].text.split('\n')
        if not paraNLines[-1].startswith("/endtemplate"):
            raise ValueError("/endtemplate muss am Anfang der letzten Zeile eines Paragraphen stehen")
        words = para0Lines[0].split()
        typ = words[1]
        if typ != "/tour" and typ != "/termin":
            raise ValueError("second word after /template must be /tour or /termin")
        typ = typ[1:]
        sel = words[2]
        if not sel.startswith("/selektion="):
            raise ValueError("third word after /template must start with /selektion=")
        sel = sel[11:].lower()
        sels = self.tourselections if typ == "tour" else self.terminselections
        if not sel in sels:
            raise ValueError("Selektion " + sel + " nicht in " + typ + "selektion")
        sel = sels[sel]
        touren = self.touren if typ == "tour" else self.termine
        self.evalTouren(sel, touren, paras)

    def evalTouren(self, sel, touren, paras):
        selectedTouren = []
        for tour in touren:
            if selektion.selected(tour, sel):
                selectedTouren.append(tour)
        if len(selectedTouren) == 0:
            return
        lastTour = selectedTouren[-1]
        for tour in selectedTouren:
            for para in paras:
                newp = insert_paragraph_copy_before(self.paraBefore, para)
                self.para = newp
                runs = newp.runs
                l = len(runs)
                for i in range(l):
                    self.runX = i
                    run = self.run = runs[i]
                    if run.text.startswith("/comment"):
                        continue
                    rtext = run.text.strip()
                    self.evalRun(run, tour)
                    if rtext == "${titel}":
                        add_hyperlink_into_run(newp, run, i, self.url)
            if tour != lastTour: # extra line between touren, not after the last one
                self.extraNl()

    def evalRun(self, run, tour):
        global debug
        if debug:
            print("run", run.text)
        linesOut = []
        linesIn = run.text.split('\n')
        for line in linesIn:
            if not line.startswith("/template") and not line.startswith("/endtemplate"):
                linesOut.append(self.expand(line, tour))
        run.text = '\n'.join(linesOut)
        self.simpleNl()

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
        return "XXBeschreibungXX"
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
        run = self.para.add_run(text="Tourleiter: ", style=self.run.style)
        run.bold = True
        move_run_before(self.runX, self.para, run)
        return ", ".join(tl)
        # self.evalLine("/bold Tourleiter: /block " + "\uaffe".join(tl), tour)

    def expAbfahrten(self, tour, format):
        afs = tour.getAbfahrten()
        if len(afs) == 0:
            return
        afl = [ af[0] + " " + af[1] + " " + af[2] for af in afs]
        return "XXAbfahrtenXX"
        #self.evalLine("/bold Ort" + ("" if len(afs) == 1 else "e") + ": /block " + "\uaffe".join(afl), tour)

    def expBetreuer(self, tour, format):
        tl = tour.getPersonen()
        if len(tl) == 0:
            return
        return "XXBetreuerXX"
        #self.evalLine("/bold Betreuer: /block " + "\uaffe".join(tl), tour)

    def expZusatzInfo(self, tour, format):
        zi = tour.getZusatzInfo()
        if len(zi) == 0:
            return
        return "XXZusatzInfoXX"
        #self.evalLine("/bold Zusatzinfo: /block " + "\uaffe".join(zi), tour)

