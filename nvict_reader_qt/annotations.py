# -*- coding: utf-8 -*-
"""Datamodellen voor tekst-annotaties en highlights.

Kleuren/lettertypes exact overgenomen uit NVict_Reader.py (color_map regel
3987-3991, font_map regel 4665-4669, kleurknoppen regel 4694-4699) zodat een
PDF die met de tkinter-app is bewerkt en met deze Qt-app wordt geopend (of
omgekeerd) dezelfde kleurwaarden gebruikt.
"""

from dataclasses import dataclass, field

FONT_MAP = {
    "Helvetica": "helv",
    "Times Roman": "tiro",
    "Courier": "cour",
}

# De "blue": (0, 0, 1)-sleutel uit de tkinter color_map is dood - de UI zette
# color_var nooit op "blue", alleen op "#0066CC" (knoplabel "Blauw"). Hier
# dus niet overgenomen.
COLOR_MAP = {
    "black": (0, 0, 0),
    "red": (1, 0, 0),
    "#0066CC": (0, 0.4, 0.8),
    "darkgreen": (0, 0.4, 0),
}

# (label, kleur-key, knop-achtergrondkleur, knop-tekstkleur) - zelfde 4 opties
# en volgorde als de tkinter-inline-editor (regel 4694-4699).
COLOR_CHOICES = [
    ("Zwart", "black", "#222222", "white"),
    ("Rood", "red", "#cc3333", "white"),
    ("Blauw", "#0066CC", "#2266bb", "white"),
    ("Groen", "darkgreen", "#228833", "white"),
]

DEFAULT_FONT_SIZE = 11
MIN_FONT_SIZE = 6
MAX_FONT_SIZE = 36

HIGHLIGHT_COLOR = (1.0, 0.85, 0.0)  # geel, zelfde als NVict_Reader.py:6498


@dataclass
class TextAnnotation:
    page_num: int
    pdf_x: float
    pdf_y: float
    text: str
    font_size: int = DEFAULT_FONT_SIZE
    color: str = "black"
    fontname: str = "helv"


@dataclass
class HighlightAnnotation:
    page_num: int
    quads: list = field(default_factory=list)  # fitz.Quad-objecten, PDF-coördinaten
    xref: int = 0  # PyMuPDF-xref van de live annotatie, voor verwijdering
