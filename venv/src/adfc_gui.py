from tkinter import *
from tkinter.filedialog import asksaveasfilename
import myLogger
import sys
import os
import tourServer
import textHandler
import printHandler
import contextlib

def toDate(dmy):  # 21.09.2018
    d = dmy[0:2]
    m = dmy[3:5]
    if len(dmy) == 10:
        y = dmy[6:10]
    else:
        y = "20" + dmy[6:8]
    if int(d) == 0 or int(d) > 31 or int(m) == 0 or int(m) > 12 or int(y) < 2000 or int(y) > 2100:
        raise ValueError("Bitte Datum als dd.mm.jjjj angeben, nicht als " + dmy)
    return y + "-" + m + "-" + d  # 2018-09-21

class TxtWriter:
    def __init__(self, targ):
        self.txt = targ

    def write(self, s):
        self.txt.insert("end", s)

class MyApp(Frame):
    def __init__(self, master):
        super().__init__(master)
        self.savFile = None
        menuBar = Menu(master)

        master.config(menu = menuBar)
        menuFile = Menu(menuBar)
        menuFile.add_command(label = "Speichern", command=self.speichern)
        menuFile.add_command(label = "Speichern unter", command=self.speichernUnter)
        menuBar.add_cascade(label="Datei", menu=menuFile)

        menuEdit = Menu(menuBar)
        menuEdit.add_command(label = "Ausschneiden", command=self.cut)
        menuEdit.add_command(label = "Kopieren", command=self.copy)
        menuEdit.add_command(label = "Einfügen", command=self.paste)
        menuBar.add_cascade(label="Bearbeiten", menu=menuEdit)

        self.createWidgets(master)

    def speichern(self):
        if self.savFile == None or self.savFile == "":
            self.fileHandler2()
            return
        with open(self.savFile, "w") as savFile:
            s = self.text.get("1.0", END)
            savFile.write(s)

    def speichernUnter(self):
        self.savFile = asksaveasfilename()
        if self.savFile == None or self.savFile == "":
            return
        self.fileHandler1()

    def cut(self):
        savedText = self.text.get(SEL_FIRST, SEL_LAST)
        self.clipboard_clear()
        self.clipboard_append(savedText)
        self.text.delete(SEL_FIRST, SEL_LAST)

    def copy(self):
        savedText = self.text.get(SEL_FIRST, SEL_LAST)
        self.clipboard_clear()
        self.clipboard_append(savedText)

    def paste(self):
        savedText = self.clipboard_get()
        if savedText is None or savedText == "":
            return
        ranges = self.text.tag_ranges(SEL)
        print("ranges=", ranges, "len ", len(ranges))
        if len(ranges) == 2:
            self.text.replace(ranges[0], ranges[1], savedText)
        else:
            self.text.insert(INSERT, savedText)

    def createWidgets(self, master):
        self.useRestVar = BooleanVar()
        self.useRestVar.set(False)
        useRestCB = Checkbutton(master, text="Aktuelle Daten werden vom Server geholt", variable=self.useRestVar)
        # print("ur={}".format(self.useRestVar.get()))

        self.usePHVar = BooleanVar()
        self.usePHVar.set(False)
        usePHCB = Checkbutton(master, text="Debug Ausgabe, ähnlich Scribus", variable=self.usePHVar)

        typen = [ "Radtour", "Termin", "Alles" ]
        typenLF = LabelFrame(master)
        typenLF["text"] = "Typen"
        self.typVar = StringVar()
        for typ in typen:
            typRB = Radiobutton(typenLF, text=typ, value=typ, variable = self.typVar)# , command = self.typHandler)
            if typ == "Alles":
                typRB.select()
            else:
                typRB.deselect()
            #typRB.pack(anchor="w")
            typRB.grid(sticky="w")

        radTypen = ["Rennrad", "Tourenrad", "Mountainbike", "Alles" ]
        radTypenLF = LabelFrame(master)
        radTypenLF["text"] = "Fahrradtyp"
        self.radTypVar = StringVar()
        for radTyp in radTypen:
            radTypRB = Radiobutton(radTypenLF, text=radTyp, value=radTyp, variable = self.radTypVar) #, command = self.radTypHandler)
            if radTyp == "Alles":
                radTypRB.select()
            else:
                radTypRB.deselect()
            #radTypRB.pack(anchor="w")
            radTypRB.grid(sticky="w")

        gliederungLB = Label(master, text="Gliederung bitte rechts eingeben:")
        self.gliederungVar = StringVar()
        self.gliederungVar.set("152085")
        gliederung = Entry(master, textvariable=self.gliederungVar)

        startDateLB = Label(master, text="Start Datum")
        self.startDateVar = StringVar()
        self.startDateVar.set("01.01.2018")
        startDate = Entry(master, textvariable=self.startDateVar)

        endDateLB = Label(master, text="Ende Datum")
        self.endDateVar = StringVar()
        self.endDateVar.set("31.12.2019")
        endDate = Entry(master, textvariable=self.endDateVar)

        startBtn = Button(master, text="Start", command=self.starten)

        textContainer = Frame(master, borderwidth=1, relief="sunken")
        self.text = Text(textContainer, wrap="none", borderwidth=0, width=100, height=50)
        textVsb = Scrollbar(textContainer, orient="vertical", command=self.text.yview)
        textHsb = Scrollbar(textContainer, orient="horizontal", command=self.text.xview)
        self.text.configure(yscrollcommand=textVsb.set, xscrollcommand=textHsb.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        textVsb.grid(row=0, column=1, sticky="ns")
        textHsb.grid(row=1, column=0, sticky="ew")
        textContainer.grid_rowconfigure(0, weight=1)
        textContainer.grid_columnconfigure(0, weight=1)

        for x in range(4):
            Grid.columnconfigure(master, x, weight=1)
        for y in range(4):
            Grid.rowconfigure(master, y, weight=1)
        useRestCB.grid(row=0, column=0, padx=10,pady=10, sticky="w")
        usePHCB.grid(row=0, column=1, padx=10,pady=10, sticky="w")
        typenLF.grid(row=1, column=0,padx=10,pady=10, sticky="w")
        radTypenLF.grid(row=1, column=1,padx=10,pady=10, sticky="w")
        gliederungLB.grid(row=2, column=0,padx=10,pady=10, sticky="w")
        gliederung.grid(row=2, column=1,padx=10,pady=10, sticky="w")
        startDateLB.grid(row=3, column=0,padx=10,pady=10, sticky="w")
        startDate.grid(row=3, column=1,padx=10,pady=10, sticky="w")
        endDateLB.grid(row=3, column=2,padx=10,pady=10, sticky="w")
        endDate.grid(row=3, column=3,padx=10,pady=10, sticky="w")
        startBtn.grid(row=4, padx=10, pady=10, sticky="w")
        textContainer.grid(row=5,columnspan = 4, padx=10,pady=10, sticky="ew")
        self.text.insert("end", "111111111111111111111111111111111111111abc11111111111111111111111111111111111111111\n")
        self.text.insert("end", "22222222222222222222222222222222222222222222222222222222222222222222222222222222\n")

    def starten(self):
        useRest = self.useRestVar.get()
        usePH = self.usePHVar.get()
        type = self.typVar.get()
        radTyp = self.radTypVar.get()
        unitKey = self.gliederungVar.get().strip()
        start = toDate(self.startDateVar.get().strip())
        end = toDate(self.endDateVar.get().strip())

        self.text.delete("1.0", END)
        txtWriter = TxtWriter(self.text)

        tourServerVar = tourServer.TourServer(False, useRest)
        if usePH:
            handler = printHandler.PrintHandler()
        else:
            handler = textHandler.TextHandler()

        touren = tourServerVar.getTouren(unitKey, start, end, type)
        with contextlib.redirect_stdout(txtWriter):
            for tour in touren:
                eventItemId = tour.get("eventItemId");
                tour = tourServerVar.getTour(eventItemId)
                if tour.isTermin():
                    handler.handleTermin(tour)
                else:
                    if radTyp != "Alles" and tour.getRadTyp() != radTyp:
                        continue
                    handler.handleTour(tour)

root = Tk()
app = MyApp(root)
app.master.title("ADFC Touren/Termine")
app.mainloop()
