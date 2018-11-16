from tkinter import *
from tkinter.filedialog import asksaveasfilename
from tkinter.filedialog import askopenfilename
from tkinter.simpledialog import askstring

# textHandler produziert output a la KV München
from myLogger import logger
import sys
import os
import tourServer
import textHandler
import rawHandler
import printHandler
import csvHandler
import pdfHandler
import contextlib
import base64
import locale

import adfc_gliederungen
from PIL import ImageTk

def toDate(dmy):  # 21.09.2018
    d = dmy[0:2]
    m = dmy[3:5]
    if len(dmy) == 10:
        y = dmy[6:10]
    else:
        y = "20" + dmy[6:8]
    if y < "2017":
        raise ValueError("Kein Datum vor 2017 möglich")
    if int(d) == 0 or int(d) > 31 or int(m) == 0 or int(m) > 12 or int(y) < 2000 or int(y) > 2100:
        raise ValueError("Bitte Datum als dd.mm.jjjj angeben, nicht als " + dmy)
    return y + "-" + m + "-" + d  # 2018-09-21

class TxtWriter:
    def __init__(self, targ):
        self.txt = targ
    def write(self, s):
        self.txt.insert("end", s)

class LabelEntry(Frame):
    def __init__(self, master, labeltext, stringtext):
        super().__init__(master)
        self.label = Label(self, text=labeltext)
        self.svar = StringVar()
        self.svar.set(stringtext)
        self.entry = Entry(self, textvariable=self.svar, width=len(stringtext) + 2, borderwidth=2)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.label.grid(row=0, column=0,sticky="w")
        self.entry.grid(row=0,column=1,sticky="w")
    def get(self):
        return self.svar.get()

class LabelOM(Frame):
    def __init__(self, master, labeltext, options, initVal, **kwargs):
        super().__init__(master)
        self.options = options
        self.label = Label(self, text=labeltext)
        self.svar = StringVar()
        self.svar.set(initVal)
        self.optionMenu = OptionMenu(self, self.svar, *options, **kwargs)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.label.grid(row=0, column=0,sticky="w")
        self.optionMenu.grid(row=0,column=1,sticky="w")
    def get(self):
        return self.svar.get()

class ListBoxSB(Frame):
    def __init__(self, master, selFunc, entries):
        super().__init__(master)
        # for the "exportselection" param see
        # https://stackoverflow.com/questions/10048609/how-to-keep-selections-highlighted-in-a-tkinter-listbox
        self.gliederungLB = Listbox(self, borderwidth=2, selectmode="extended", exportselection=False, width=50)
        self.gliederungLB.bind("<<ListboxSelect>>", selFunc)

        self.entries = sorted(entries)
        self.gliederungLB.insert("end", *self.entries)
        self.lbVsb = Scrollbar(self, orient="vertical", command=self.gliederungLB.yview)
        self.gliederungLB.configure(yscrollcommand=self.lbVsb.set)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)
        self.gliederungLB.grid(row=0, column=0,sticky="nsew")
        self.lbVsb.grid(row=0,column=1,sticky="ns")
    def curselection(self):
        names = [ self.entries[i] for i in self.gliederungLB.curselection() ]
        if "0 Alles" in names:
            return "Alles"
        s = ",".join( [name.split(maxsplit=1)[0] for name in names] )
        return s
    def clearLB(self):
        self.gliederungLB.selection_clear(0,len(self.entries))
    def setEntries(self, entries):
        self.entries = sorted(entries)
        self.gliederungLB.delete(0, "end")
        self.gliederungLB.insert("end", *self.entries)


class MyApp(Frame):
    def __init__(self, master):
        super().__init__(master)
        self.savFile = None
        self.pos = None
        self.searchVal = ""
        self.images = []
        self.pdfTemplateName = ""
        menuBar = Menu(master)
        master.config(menu = menuBar)
        menuFile = Menu(menuBar)
        menuFile.add_command(label = "Speichern", command=self.store, accelerator="Ctrl+s")
        master.bind_all("<Control-s>", self.store)
        menuFile.add_command(label = "Speichern unter", command=self.storeas)
        menuFile.add_command(label = "PDF Template", command=self.pdfTemplate)
        menuBar.add_cascade(label = "Datei", menu=menuFile)

        menuEdit = Menu(menuBar)
        menuEdit.add_command(label = "Ausschneiden", command=self.cut, accelerator="Ctrl+x")
        master.bind_all("<Control-x>", self.cut)
        menuEdit.add_command(label = "Kopieren", command=self.copy, accelerator="Ctrl+c")
        master.bind_all("<Control-c>", self.copy)
        menuEdit.add_command(label = "Einfügen", command=self.paste, accelerator="Ctrl+v")
        master.bind_all("<Control-v>", self.paste)
        menuEdit.add_command(label = "Suchen", command=self.search, accelerator="Ctrl+f")
        master.bind_all("<Control-f>", self.search)
        menuEdit.add_command(label = "Erneut suchen", command=self.searchAgain, accelerator="F3")
        master.bind_all("<F3>", self.searchAgain)
        menuBar.add_cascade(label = "Bearbeiten", menu=menuEdit)
        self.createWidgets(master)

    def createPhoto(self, b64):
        binary = base64.decodebytes(b64.encode())
        photo = ImageTk.PhotoImage(data=binary)
        return photo

    def store(self, *args):
        if self.savFile == None or self.savFile == "":
            self.storeas()
            return
        with open(self.savFile, "w", encoding="utf-8-sig") as savFile:
            s = self.text.get("1.0", END)
            savFile.write(s)

    def storeas(self, *args):
        format = self.formatOM.get()
        self.savFile = asksaveasfilename(title="Select file", initialfile="adfc_export",
            defaultextension=".csv" if format == "CSV" else ".txt",
            filetypes=[("CSV", ".csv")] if format == "CSV" else [("TXT", ".txt")])
        if self.savFile == None or self.savFile == "":
            return
        self.store()

    def pdfTemplate(self, *args):
        self.pdfTemplateName = askopenfilename(title="Choose a PDF Template",
            defaultextension=".json",
            filetypes=[("JSON", ".json")])

    def cut(self, *args):
        savedText = self.text.get(SEL_FIRST, SEL_LAST)
        self.clipboard_clear()
        self.clipboard_append(savedText)
        self.text.delete(SEL_FIRST, SEL_LAST)

    def copy(self, *args):
        savedText = self.text.get(SEL_FIRST, SEL_LAST)
        self.clipboard_clear()
        self.clipboard_append(savedText)

    def paste(self, *args):
        savedText = self.clipboard_get()
        if savedText is None or savedText == "":
            return
        ranges = self.text.tag_ranges(SEL)
        if len(ranges) == 2:
            self.text.replace(ranges.first, ranges.last, savedText)
        else:
            self.text.insert(INSERT, savedText)

    def search(self, *args):
        self.searchVal = askstring("Suchen", "Bitte Suchstring eingeben", initialvalue=self.searchVal)
        if self.searchVal is None:
            return
        self.searchAgain()

    def searchAgain(self, *args):
        self.pos = self.text.search(self.searchVal, INSERT + "+1c", END)
        if self.pos != "":
            self.text.mark_set(INSERT, self.pos)
            self.text.see(self.pos)
        self.text.focus_set()

    def typHandler(self):
        typ = self.typVar.get()
        for rtBtn in self.radTypBtns:
            if typ == "Termin":
                rtBtn.config(state=DISABLED)
            else:
                rtBtn.config(state=NORMAL)

    def gliederungSel(self, event):
        sel = self.gliederungLBSB.curselection()
        self.gliederungSvar.set(sel)

    def clearLB(self, event):
        self.gliederungLBSB.clearLB()

    def lvSelector(self, event):
        kvMap = adfc_gliederungen.getLV(event[0:3])
        entries = [ key + " " + kvMap[key] for key in kvMap.keys() ]
        self.gliederungLBSB.setEntries(entries)
        self.gliederungSvar.set("")

    def createWidgets(self, master):
        self.useRestVar = BooleanVar()
        self.useRestVar.set(False)
        useRestCB = Checkbutton(master, text="Aktuelle Daten werden vom Server geholt", variable=self.useRestVar)

        self.includeSubVar = BooleanVar()
        self.includeSubVar.set(True)
        includeSubCB = Checkbutton(master, text="Untergliederungen einbeziehen", variable=self.includeSubVar)

        self.formatOM = LabelOM(master, "Ausgabeformat:", ["München", "Starnberg", "CSV", "Text", "PDF"], "PDF")
        self.linkTypeOM = LabelOM(master, "Links to:", ["Frontend", "Backend", "Keine"], "frontEnd")

        typen = [ "Radtour", "Termin", "Alles" ]
        typenLF = LabelFrame(master)
        typenLF["text"] = "Typen"
        self.typVar = StringVar()
        self.typBtns = []
        for typ in typen:
            typRB = Radiobutton(typenLF, text=typ, value=typ, variable = self.typVar, command = self.typHandler)
            self.typBtns.append(typRB)
            if typ == "Alles":
                typRB.select()
            else:
                typRB.deselect()
            typRB.grid(sticky="w")

        radTypen = ["Rennrad", "Tourenrad", "Mountainbike", "Alles" ]
        radTypenLF = LabelFrame(master)
        radTypenLF["text"] = "Fahrradtyp"
        self.radTypVar = StringVar()
        self.radTypBtns = []
        for radTyp in radTypen:
            radTypRB = Radiobutton(radTypenLF, text=radTyp, value=radTyp, variable = self.radTypVar) #, command = self.radTypHandler)
            self.radTypBtns.append(radTypRB)
            if radTyp == "Alles":
                radTypRB.select()
            else:
                radTypRB.deselect()
            radTypRB.grid(sticky="w")


        # container for LV selector and Listbox for KVs
        glContainer = Frame(master, borderwidth=2, relief="sunken", width=100)
        # need a tourServer here early for list of LVs
        tourServerVar = tourServer.TourServer(False, True, False)
        lvMap = adfc_gliederungen.getLVs()
        self.lvList = [ key + " " + lvMap[key] for key in lvMap.keys() ]
        self.lvList = sorted(self.lvList)
        self.lvOM = LabelOM(glContainer, "Landesverband:", self.lvList, "152", command=self.lvSelector)
        kvMap = adfc_gliederungen.getLV(self.lvOM.get()[0:3])
        entries = [ key + " " + kvMap[key] for key in kvMap.keys() ]
        self.gliederungLBSB = ListBoxSB(glContainer, self.gliederungSel, entries)
        self.gliederungSvar = StringVar()
        self.gliederungSvar.set("152085")
        self.gliederungEN = Entry(master, textvariable=self.gliederungSvar, borderwidth=2, width=60)
        self.gliederungEN.bind("<Key>", self.clearLB)
        self.lvOM.grid(row=0, column=0, sticky="nsew")
        self.gliederungLBSB.grid(row=1, column=0, sticky="nsew")
        glContainer.grid_rowconfigure(0, weight=1)
        glContainer.grid_rowconfigure(1, weight=1)
        glContainer.grid_columnconfigure(0, weight=1)

        self.startDateLE = LabelEntry(master, "Start Datum:", "01.01.2018")
        self.endDateLE = LabelEntry(master, "Ende Datum:", "31.12.2019")
        startBtn = Button(master, text="Start", bg="red", command=self.starten)

        textContainer = Frame(master, borderwidth=2, relief="sunken")
        self.text = Text(textContainer, wrap="none", borderwidth=0, cursor="arrow") # width=100, height=40,
        textVsb = Scrollbar(textContainer, orient="vertical", command=self.text.yview)
        textHsb = Scrollbar(textContainer, orient="horizontal", command=self.text.xview)
        self.text.configure(yscrollcommand=textVsb.set, xscrollcommand=textHsb.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        textVsb.grid(row=0, column=1, sticky="ns")
        textHsb.grid(row=1, column=0, sticky="ew")
        textContainer.grid_rowconfigure(0, weight=1)
        textContainer.grid_columnconfigure(0, weight=1)

        for x in range(2):
            Grid.columnconfigure(master, x, weight= 1 if x == 1 else 0)
        for y in range(7):
            Grid.rowconfigure(master, y, weight= 1 if y == 6 else 0)
        useRestCB.grid(row=0, column=0, padx=5,pady=2, sticky="w")
        includeSubCB.grid(row=0, column=1, padx=5,pady=2, sticky="w")
        self.formatOM.grid(row=1, column=0, padx=5,pady=2, sticky="w")
        self.linkTypeOM.grid(row=1, column=1, padx=5,pady=2, sticky="w")
        typenLF.grid(row=2, column=0,padx=5,pady=2, sticky="w")
        radTypenLF.grid(row=2, column=1,padx=5,pady=2, sticky="w")
        glContainer.grid(row=3, column=0,padx=5,pady=2, sticky="w")
        self.gliederungEN.grid(row=3, column=1,padx=5,pady=2, sticky="w")
        self.startDateLE.grid(row=4, column=0,padx=5,pady=2, sticky="w")
        self.endDateLE.grid(row=4, column=1,padx=5,pady=2, sticky="w")

        startBtn.grid(row=5, padx=5, pady=2, sticky="w")
        textContainer.grid(row=6,columnspan = 2, padx=5,pady=2, sticky="nsew")

        self.pos = "1.0"
        self.text.mark_set(INSERT, self.pos)

    def insertImage(self, tour):
        img = tour.getImagePreview()
        if img != None:
            print()
            photo = self.createPhoto(img)
            self.images.append(photo)  # see http://effbot.org/pyfaq/why-do-my-tkinter-images-not-appear.htm
            self.text.image_create(INSERT, image=photo)
            print()

    def getLinkType(self):
        return self.linkTypeOM.get().lower()
    def getRadTyp(self):
        return self.radTypVar.get()
    def getTyp(self):
        return self.typVar.get()
    def getGliederung(self):
        return self.gliederungSvar.get()
    def getIncludeSub(self):
        return self.includeSubVar.get()
    def getStart(self):
        return self.startDateLE.get().strip()
    def getEnd(self):
        return self.endDateLE.get().strip()

    def starten(self):
        useRest = self.useRestVar.get()
        includeSub = self.includeSubVar.get()
        type = self.typVar.get()
        radTyp = self.radTypVar.get()
        unitKeys = self.gliederungSvar.get().split(",")
        start = toDate(self.startDateLE.get().strip())
        end = toDate(self.endDateLE.get().strip())
        self.images.clear()

        self.text.delete("1.0", END)
        txtWriter = TxtWriter(self.text)

        formatS = self.formatOM.get()
        if formatS == "Starnberg":
            handler = printHandler.PrintHandler()
        elif formatS == "München":
            handler = textHandler.TextHandler()
        elif formatS == "CSV":
            handler = csvHandler.CsvHandler(txtWriter)
        elif formatS == "Text":
            handler = rawHandler.RawHandler()
        elif formatS == "PDF":
            handler = pdfHandler.PDFHandler(self)
            # conditions obtained from PDF template!
            includeSub = handler.getIncludeSub()
            type = handler.getType()
            radTyp = handler.getRadTyp()
            unitKeys = handler.getUnitKeys().split(",")
            start = toDate(handler.getStart())
            end = toDate(handler.getEnd())
        else:
            handler = rawHandler.RawHandler()

        tourServerVar = tourServer.TourServer(False, useRest, includeSub)
        touren = []
        for unitKey in unitKeys:
            if unitKey == "Alles":
                unitKey = ""
            touren.extend(tourServerVar.getTouren(unitKey.strip(), start, end, type))

        if (isinstance(handler, textHandler.TextHandler)
            or isinstance(handler, csvHandler.CsvHandler)
            or isinstance(handler, pdfHandler.PDFHandler)
            or isinstance(handler, rawHandler.RawHandler)):
            tourServerVar.calcNummern()

        def tourdate(self):
            return self.get("beginning")
        touren.sort(key=tourdate)  # sortieren nach Datum

        with contextlib.redirect_stdout(txtWriter):
            if len(touren) == 0:
                handler.nothingFound()
            for tour in touren:
                tour = tourServerVar.getTour(tour)
                if tour.isTermin():
                    if isinstance(handler, rawHandler.RawHandler):
                        self.insertImage(tour)
                    handler.handleTermin(tour)
                else:
                    if radTyp != "Alles" and tour.getRadTyp() != radTyp:
                        continue
                    if isinstance(handler, rawHandler.RawHandler):
                        self.insertImage(tour)
                    handler.handleTour(tour)
            if isinstance(handler, pdfHandler.PDFHandler): # TODO
                handler.handleEnd()
        self.pos = "1.0"
        self.text.mark_set(INSERT, self.pos)
        self.text.focus_set()

#locale.setlocale(locale.LC_ALL, "de_DE")
locale.setlocale(locale.LC_TIME, "German")
root = Tk()
app = MyApp(root)
app.master.title("ADFC Touren/Termine")
app.mainloop()


"""
TODO:
delete adfc_rest2?
parse PDF, expand template within PDF document???
create docx instead of pdf?
"""
