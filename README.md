# adfc2
read bicycle events from "ADFC Tourenportal" and convert to text 
or Word (.docx) or something else.

This is a Python program either running in a UI, if started from adfc_gui.py,
or running as a Scribus script, when Scribus is told to run the script 
adfc_rest2.py or scrbHandler.py, or running as a cmdline program.

It searches first for all tours belonging to a "Gliederung", i.e. ADFC
sub organization, then gets detailed info about each tour, and outputs
them in various formats, including MSWord/.docx, or, when called as a
Scribus script, controls Scribus output, or produces other output.

For adfc_rest2.py Scribus requires a document with some predefined
styles like Radtour_Title. For scrbHandler.py Scribus needs a
document with a parameter and one or more template sections.

When tp2vadb.py is called, the program creates either an XML file 
that is used to publish events to the "Veranstaltungsdatenbank Hamburg",
or CALDAV entries.

To create a leaflet containing event data I recommend 
to use Affinity Publisher and import the Word file. 
You may also use Scribus since Version 1.5.6. 
Older versions use Python2, and do not support getCharacterStyle().

