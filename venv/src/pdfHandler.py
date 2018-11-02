# encoding: utf-8

import tourRest
import os
import json
import functools
from myLogger import logger

from myLogger import logger
from fpdf import FPDF

schwierigkeitMap = { 0: "sehr einfach", 1: "sehr einfach", 2: "einfach", 3: "mittel", 4: "schwer", 5: "sehr schwer"}

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

class PDFHandler:
    def __init__(self, gui):
        self.gui = gui
        self.pdf = FPDF('P', "mm", "A4")
        self.pdf.add_page()
        self.pdf.set_auto_page_break(True)
        self.styles = {}
        self.terminselections = {}
        self.tourselections = {}
        self.touren = []
        self.termine = []
        self.gui.pdfTemplateName = "C:/Users/Michael/PycharmProjects/ADFC1/venv/src/template.json" # TODO
        if self.gui.pdfTemplateName is None or self.gui.pdfTemplateName == "":
            self.gui.pdfTemplate()
        if self.gui.pdfTemplateName is None or self.gui.pdfTemplateName == "":
            raise ValueError("must specify path to PDF template!")
        with open(self.gui.pdfTemplateName, "r", encoding="utf8") as jsonFile:
            self.pdfJS = json.load(jsonFile)
        self.parseTemplate()

    def nothingFound(self):
        logger.info("Nichts gefunden")
        print("Nichts gefunden")

    def parseTemplate(self):
        for key in ["margins", "fonts", "styles", "selection", "text"]:
            if self.pdfJS.get(key) == None:
                raise ValueError("pdf template " + self.gui.pdfTemplateName + " must have a section " + key)
        margins = self.pdfJS.get("margins")
        leftMargin = margins.get("left")
        rightMargin = margins.get("right")
        topMargin = margins.get("top")
        self.pdf.set_margins(left=leftMargin, top=topMargin, right=rightMargin)

        self.margins = (leftMargin, topMargin, rightMargin)
        self.spacing = margins.get("spacing") # float
        fonts = self.pdfJS.get("fonts")
        for font in iter(fonts):
            family = font.get("family")
            if family is None or family == "":
                raise ValueError("font family not specified")
            file = font.get("file")
            if file is None or file == "":
                raise ValueError("font file not specified")
            fontStyle = font.get("fontStyle")
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
        self.includeSub = selection.get("includeSub")
        self.start = selection.get("start")
        self.end = selection.get("end")
        sels = selection.get("terminselection")
        for sel in iter(sels):
            self.terminselections[sel.get("name")] = sel
        sels = selection.get("tourselection")
        for sel in iter(sels):
            self.tourselections[sel.get("name")] = sel

    def getIncludeSub(self):
        return self.includeSub
    def getType(self):
        if len(self.terminselections) != 0 and len(self.tourselections) != 0:
            return "Alles";
        if len(self.terminselections) != 0:
            return "Termin"
        if len(self.tourselections) != 0:
            return "Radtour"
    def getRadTyp(self):
        s = set()
        for sel in self.tourselections.values():
            s.add(sel.get("radtyp"))
        if "Alles" in s:
            return "Alles";
        if len(s) == 1:
            return s[0]
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
                self.evalLine(line)
                lineNo += 1
        self.pdf.output(dest='F', name="testfpdf.pdf")

    def evalLine(self, line):
        print("line", line)
        text = []
        self.align = "L"
        self.fontstyles = ""
        self.curStyle = self.styles.get("body")
        words = line.split(' ')
        for i in range(len(words)):
            word = words[i];
            if word.startswith("/"):
                cmd = word[1:]
                if cmd in self.styles.keys():
                    self.handleText(' '.join(text))
                    text = []
                    self.curStyle = self.styles.get(cmd)
                elif cmd in ["right", "left", "center", "block"]:
                    self.handleText(' '.join(text))
                    text = []
                    self.align = cmd[0].upper()
                    if self.align == 'B':
                        self.align = 'J' ## justification
                elif cmd in ["bold", "italic", "underline"]:
                    self.handleText(' '.join(text))
                    text = []
                    self.fontstyles += cmd[0].upper()
                else:
                    text.append(word)
            else:
                text.append(word)
        self.handleText(' '.join(text))
        if not self.align == "J": # multicell does newline automatically
            y = self.pdf.get_y()
            self.pdf.ln()
            y = self.pdf.get_y()

    def evalTemplate(self, lines):
        print("template:")
        for line in lines:
            if line.startswith("/comment"):
                continue
            self.evalLine(line)

    def handleText(self, s:str):
        s = s.strip()
        if s == "":
            return
        print("Text:", s)
        s += " "
        if self.curStyle.type == "image":
            self.drawImage(s)
            return
        style = self.curStyle.copy()
        for fs in self.fontstyles:
            if style.fontStyle.find(fs) == -1:
                style.fontStyle += fs
        self.fontstyles = ""
        self.setStyle(style)
        h=(style.size * 0.35278 + self.spacing)
        if self.align == 'J':
            self.pdf.multi_cell(w=0, h=h, txt=s, border=0, align=self.align, fill=0)
        elif self.align == 'R':
            self.pdf.cell(w=0, h=h, txt=s, border=0, ln=0, align=self.align, fill=0)
        else:
            w = self.pdf.get_string_width(s)
            x = self.pdf.get_x()
            y = self.pdf.get_y() #TODO
            if (x + w) >= (210 - self.margins[2]): # i.e. exceeds right margin
                # self.pdf.multi_cell(w=0, h=h, txt=s, border=0, align=self.align, fill=0)
                self.multiline(h, s)
            else:
                self.pdf.cell(w=w, h=h, txt=s, border=0, ln=0, align=self.align, fill=0)

    def multiline(self, h: float, s:str):
        """ line too long, see if I can split line after blank """
        x = self.pdf.get_x()
        l = len(s)
        while l > 0:
            lb = s.rfind(' ', 0, l) # last blank
            if lb == -1: # can not split line
                if x > self.margins[0]:
                    self.pdf.ln()
                w = self.pdf.get_string_width(s)
                self.pdf.cell(w=w, h=h, txt=s, border=0, ln=0, align=self.align, fill=0)
                return
            sub = s[0:lb]
            w = self.pdf.get_string_width(sub)
            if (x + w) >= (210 - self.margins[2]):
                l = lb
                continue
            self.pdf.cell(w=w, h=h, txt=sub, border=0, ln=1, align=self.align, fill=0)
            x = self.pdf.get_x()
            s = s[lb + 1:]
            l = len(s)


    def setStyle(self, style:Style):
        print("Style:", style)
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
            x = 210 - self.margins[2] - w - 10
        y -= h  # align lower edge of image with baseline of text (or so)
        if y < self.margins[1]:
            y = self.margins[1]
        self.pdf.image(imgName.strip(), x=x, y=y, w=w) # h=h
        self.pdf.set_y(self.pdf.get_y() + 7)
