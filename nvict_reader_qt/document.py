# -*- coding: utf-8 -*-
"""PyMuPDF (fitz) lazy import, overgenomen uit NVict_Reader.py (regel 97-103).

PySide6 zelf kan niet lazy geladen worden (de UI heeft het meteen nodig),
maar fitz pas bij het daadwerkelijk openen van een PDF houdt de opstarttijd
laag, net als in de tkinter-versie.
"""

_fitz = None


def get_fitz():
    """Lazy import van PyMuPDF - alleen laden wanneer een PDF wordt geopend.

    Importeert als `pymupdf` (niet de oude `fitz`-naam): functioneel
    identiek, maar `import fitz` drukt bij elke start een deprecation-
    warning af ("gebruik pymupdf i.p.v. fitz") die de gebruiker bij het
    handmatig starten van het programma te zien kreeg.
    """
    global _fitz
    if _fitz is None:
        import pymupdf
        _fitz = pymupdf
    return _fitz
