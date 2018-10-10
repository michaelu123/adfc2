# adfc2
read bicycle tours from "ADFC Tourenportal" and convert to text

This is a Python program either running in a UI, if started from adfc_gui.py,
or as a command line program, if started from adfc_rest2.py, or running as a
Scribus script, when Scribus is told to run the script adfc_rest2.py.

It searches first for all tours belonging to a "Gliederung", i.e. ADFC
sub organization, then gets detailed info about each tour, and outputs
them in various formats.
