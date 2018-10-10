from tkinter import *
from tkinter.filedialog import asksaveasfilename
from tkinter.simpledialog import askstring

# textHandler produziert output a la KV München
import myLogger
import sys
import os
import tourServer
import textHandler
import printHandler
import csvHandler
import contextlib
import base64
import adfc_gliederungen
from PIL import ImageTk


name2num = {} # dict maps name to num
num2name = {} # dict maps num to name


def toDate(dmy):  # 21.09.2018
    d = dmy[0:2]
    m = dmy[3:5]
    if len(dmy) == 10:
        y = dmy[6:10]
    else:
        y = "20" + dmy[6:8]
    if y < "2018":
        raise ValueError("Kein Datum vor 2018 möglich")
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
    def __init__(self, master, labeltext, options):
        super().__init__(master)
        self.options = options
        self.label = Label(self, text=labeltext)
        self.svar = StringVar()
        self.svar.set(options[0])
        self.optionMenu = OptionMenu(self, self.svar, *options)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.label.grid(row=0, column=0,sticky="w")
        self.optionMenu.grid(row=0,column=1,sticky="w")
    def get(self):
        return self.svar.get()

class ListBoxSB(Frame):
    def __init__(self, master, selFunc):
        super().__init__(master)
        # for the "exportselection" param see
        # https://stackoverflow.com/questions/10048609/how-to-keep-selections-highlighted-in-a-tkinter-listbox
        self.gliederungLB = Listbox(self, borderwidth=2, selectmode="extended", exportselection=False)
        self.gliederungLB.bind("<<ListboxSelect>>", selFunc)
        self.entries = num2name.values()
        self.entries = sorted(self.entries)
        self.gliederungLB.insert("end", *self.entries)
        self.lbVsb = Scrollbar(self, orient="vertical", command=self.gliederungLB.yview)
        self.gliederungLB.configure(yscrollcommand=self.lbVsb.set)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.gliederungLB.grid(row=0, column=0,sticky="w")
        self.lbVsb.grid(row=0,column=1,sticky="ns")
    def curselection(self):
        names = [ self.entries[i] for i in self.gliederungLB.curselection() ]
        if "Alle" in names:
            return "Alle"
        s = ",".join( [name2num[name] for name in names] )
        return s
    def clearLB(self):
        self.gliederungLB.selection_clear(0,len(self.entries))

class MyApp(Frame):
    def __init__(self, master):
        super().__init__(master)
        self.savFile = None
        self.pos = None
        self.searchVal = ""
        self.images = []
        menuBar = Menu(master)
        master.config(menu = menuBar)
        menuFile = Menu(menuBar)
        menuFile.add_command(label = "Speichern", command=self.store, accelerator="Ctrl+s")
        master.bind_all("<Control-s>", self.store)
        menuFile.add_command(label = "Speichern unter", command=self.storeas)
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

        global num2name, name2num
        num2name = adfc_gliederungen.gliederungDict
        name2num = {num2name.get(k):k for k in num2name.keys() }
        self.createWidgets(master)

    def createPhoto(self, b64):
        binary = base64.decodebytes(b64.encode())
        photo = ImageTk.PhotoImage(data=binary)
        return photo

    def store(self, *args):
        if self.savFile == None or self.savFile == "":
            self.storeas()
            return
        with open(self.savFile, "w") as savFile:
            s = self.text.get("1.0", END)
            savFile.write(s)

    def storeas(self, *args):
        self.savFile = asksaveasfilename()
        if self.savFile == None or self.savFile == "":
            return
        self.store()

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
        sel = self.gliederungLB.curselection()
        self.gliederungSvar.set(sel)

    def clearLB(self, event):
        self.gliederungLB.clearLB()

    def createWidgets(self, master):
        self.useRestVar = BooleanVar()
        self.useRestVar.set(False)
        useRestCB = Checkbutton(master, text="Aktuelle Daten werden vom Server geholt", variable=self.useRestVar)
        # print("ur={}".format(self.useRestVar.get()))

        self.formatOM = LabelOM(master, "AusgabeFormat:", ["Starnberg", "München", "CSV"])

        typen = [ "Radtour", "Termin", "Alles" ]
        typenLF = LabelFrame(master)
        typenLF["text"] = "Typen"
        self.typVar = StringVar()
        for typ in typen:
            typRB = Radiobutton(typenLF, text=typ, value=typ, variable = self.typVar, command = self.typHandler)
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

        self.gliederungLB = ListBoxSB(master, self.gliederungSel)
        self.gliederungSvar = StringVar()
        self.gliederungSvar.set("152085,15208501,15208507,15208512,15208514")
        self.gliederungEN = Entry(master, textvariable=self.gliederungSvar, borderwidth=2, width=60)
        self.gliederungEN.bind("<Key>", self.clearLB)

        self.startDateLE = LabelEntry(master, "Start Datum:", "01.01.2018")
        self.endDateLE = LabelEntry(master, "Ende Datum:", "31.12.2019")
        startBtn = Button(master, text="Start", command=self.starten)

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
        for y in range(6):
            Grid.rowconfigure(master, y, weight= 1 if y == 5 else 0)
        useRestCB.grid(row=0, column=0, padx=5,pady=5, sticky="w")
        self.formatOM.grid(row=0, column=1, padx=5,pady=5, sticky="w")
        typenLF.grid(row=1, column=0,padx=5,pady=5, sticky="w")
        radTypenLF.grid(row=1, column=1,padx=5,pady=5, sticky="w")
        self.gliederungLB.grid(row=2, column=0,padx=5,pady=5, sticky="w")
        self.gliederungEN.grid(row=2, column=1,padx=5,pady=5, sticky="w")
        self.startDateLE.grid(row=3, column=0,padx=5,pady=5, sticky="w")
        self.endDateLE.grid(row=3, column=1,padx=5,pady=5, sticky="w")

        startBtn.grid(row=4, padx=5, pady=5, sticky="w")
        textContainer.grid(row=5,columnspan = 2, padx=5,pady=5, sticky="nsew")

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

    def starten(self):
        useRest = self.useRestVar.get()
        type = self.typVar.get()
        radTyp = self.radTypVar.get()
        unitKeys = self.gliederungSvar.get().split(",")
        start = toDate(self.startDateLE.get().strip())
        end = toDate(self.endDateLE.get().strip())
        self.images.clear()

        self.text.delete("1.0", END)
        txtWriter = TxtWriter(self.text)

        tourServerVar = tourServer.TourServer(False, useRest)
        formatS = self.formatOM.get()
        if formatS == "Starnberg":
            handler = printHandler.PrintHandler()
        elif formatS == "München":
            handler = textHandler.TextHandler()
        elif formatS == "CSV":
            handler = csvHandler.CsvHandler(txtWriter)
        else:
            handler = textHandler.TextHandler()

        touren = []
        for unitKey in unitKeys:
            if unitKey == "Alle":
                unitKey = ""
            touren.extend(tourServerVar.getTouren(unitKey.strip(), start, end, type, isinstance(handler, textHandler.TextHandler)))

        if (isinstance(handler, textHandler.TextHandler) or isinstance(handler, csvHandler.CsvHandler))\
                and (type == "Radtour" or type == "Alles"):
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
                    if isinstance(handler, textHandler.TextHandler):
                        self.insertImage(tour)
                    handler.handleTermin(tour)
                else:
                    if radTyp != "Alles" and tour.getRadTyp() != radTyp:
                        continue
                    if isinstance(handler, textHandler.TextHandler):
                        self.insertImage(tour)
                    handler.handleTour(tour)
        self.pos = "1.0"
        self.text.mark_set(INSERT, self.pos)
        self.text.focus_set()

root = Tk()
app = MyApp(root)
app.master.title("ADFC Touren/Termine")
app.mainloop()
