from myLogger import logger

def selTitelEnthält(tour, lst):
    titel = tour.getTitel()
    for  elem in lst:
        if titel.find(elem) >= 0:
            return True
    return False

def selTitelEnthältNicht(tour, lst):
    titel = tour.getTitel()
    for  elem in lst:
        if titel.find(elem) >= 0:
            return False
    return True

def selRadTyp(tour, lst):
    if "Alles" in lst:
        return True
    radTyp = tour.getRadTyp()
    return radTyp in lst

def selTourNr(tour, lst):
    nr = int(tour.getNummer())
    return nr in lst

def selNotTourNr(tour, lst):
    nr = int(tour.getNummer())
    return not nr in lst

def selKategorie(tour, lst):
    kat = tour.getKategorie()
    return kat in lst

def selMerkmalEnthält(tour, lst):
    merkmale = tour.getMerkmale()
    for merkmal in merkmale:
        for val in lst:
            if merkmal.find(val) >= 0:
                return True
    return False

def selMerkmalEnthältNicht(tour, lst):
    merkmale = tour.getMerkmale()
    for merkmal in merkmale:
        for val in lst:
            if merkmal.find(val) >= 0:
                return False
    return True

def selected(tour, sel):
    for key in sel.keys():
        if key == "name" or key.startswith("comment"):
            continue
        try:
            f = selFunctions[key]
            lst = sel[key]
            if not f(tour, lst):
                return False
        except Exception:
            logger.exception("no function for selection verb " + key + " in selection " + sel.get("name"))
    else:
        return True

selFunctions = {

    "titelenthält": selTitelEnthält,
    "titelenthältnicht": selTitelEnthältNicht,
    "terminnr": selTourNr,
    "nichtterminnr": selNotTourNr,
    "tournr": selTourNr,
    "nichttournr": selNotTourNr,
    "radtyp": selRadTyp,
    "kategorie": selKategorie,
    "merkmalenthält": selMerkmalEnthält,
    "merkmalenthältnicht": selMerkmalEnthältNicht
}

def getSelFunctions():
    return selFunctions


