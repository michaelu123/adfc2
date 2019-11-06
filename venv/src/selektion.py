# encoding: utf-8
from myLogger import logger

def selTitelEnthaelt(event, lst):
    titel = event.getTitel()
    for  elem in lst:
        if titel.find(elem) >= 0:
            return True
    logger.debug("event %s nicht enthält %s", event.getTitel(), lst)
    return False

def selTitelEnthaeltNicht(event, lst):
    titel = event.getTitel()
    for  elem in lst:
        if titel.find(elem) >= 0:
            logger.debug("event %s nicht enthältnicht %s", event.getTitel(), elem)
            return False
    return True

def selRadTyp(tour, lst):
    if "Alles" in lst:
        return True
    radTyp = tour.getRadTyp()
    if radTyp in lst:
        return True
    logger.debug("tour %s nicht radtyp %s", tour.getTitel(), lst)
    return False

def selTourNr(tour, lst):
    nr = int(tour.getNummer())
    if nr in lst:
        return True
    logger.debug("tour %s nicht tournr %s", tour.getTitel(), lst)
    return False

def selNotTourNr(tour, lst):
    nr = int(tour.getNummer())
    if nr in lst:
        logger.debug("tour %s nicht nichttournr %s", tour.getTitel(), lst)
        return False
    else:
        return True

def selKategorie(event, lst):
    kat = event.getKategorie()
    if kat in lst:
        return True
    logger.debug("event %s nicht kategorie %s", event.getTitel(), lst)
    return False

def selMerkmalEnthaelt(event, lst):
    merkmale = event.getMerkmale()
    for merkmal in merkmale:
        for val in lst:
            if merkmal.find(val) >= 0:
                return True
    logger.debug("event %s nicht merkmale %s in %s", event.getTitel(), merkmale, lst)
    return False

def selMerkmalEnthaeltNicht(event, lst):
    merkmale = event.getMerkmale()
    for merkmal in merkmale:
        for val in lst:
            if merkmal.find(val) >= 0:
                logger.debug("event %s nicht nichtmerkmale %s in %s", event.getTitel(), merkmale, lst)
                return False
    return True

def selected(event, sel):
    for key in sel.keys():
        if key == "name" or key.startswith("comment"):
            continue
        try:
            f = selFunctions[key]
            lst = sel[key]
            if not f(event, lst):
                return False
        except Exception:
            logger.exception("Keine Funktion für den Ausdruck " + key + " in der Selektion " + sel.get("name") + " gefunden")
    else:
        logger.debug("event %s selected", event.getTitel())
        return True

selFunctions = {

    "titelenthält": selTitelEnthaelt,
    "titelenthältnicht": selTitelEnthaeltNicht,
    "terminnr": selTourNr,
    "nichtterminnr": selNotTourNr,
    "tournr": selTourNr,
    "nichttournr": selNotTourNr,
    "radtyp": selRadTyp,
    "kategorie": selKategorie,
    "merkmalenthält": selMerkmalEnthaelt,
    "merkmalenthältnicht": selMerkmalEnthaeltNicht
}

def getSelFunctions():
    return selFunctions


