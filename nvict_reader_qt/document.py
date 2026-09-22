# -*- coding: utf-8 -*-
"""PyMuPDF (fitz) lazy import, overgenomen uit NVict_Reader.py (regel 97-103).

PySide6 zelf kan niet lazy geladen worden (de UI heeft het meteen nodig),
maar fitz pas bij het daadwerkelijk openen van een PDF houdt de opstarttijd
laag, net als in de tkinter-versie.
"""

_fitz = None


def get_fitz():
    """Lazy import van PyMuPDF (fitz) - alleen laden wanneer een PDF wordt geopend."""
    global _fitz
    if _fitz is None:
        import fitz
        _fitz = fitz
    return _fitz
