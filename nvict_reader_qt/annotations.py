# -*- coding: utf-8 -*-
"""Datamodellen voor tekst-annotaties en highlights.

Lettertypes overgenomen uit NVict_Reader.py (font_map regel 4665-4669).
Kleur is een vrij te kiezen hex-string (via QColorDialog, testfeedback
fase 9) i.p.v. tkinter's 4 vaste voorkeurskleuren.
"""

from dataclasses import dataclass, field

FONT_MAP = {
    "Helvetica": "helv",
    "Times Roman": "tiro",
    "Courier": "cour",
}

DEFAULT_FONT_SIZE = 11
MIN_FONT_SIZE = 6
MAX_FONT_SIZE = 36
DEFAULT_TEXT_COLOR = "#000000"

HIGHLIGHT_COLOR = (1.0, 0.85, 0.0)  # geel, zelfde als NVict_Reader.py:6498


def hex_to_fitz_rgb(hex_color: str):
    """Zet '#RRGGBB' om naar een (r, g, b)-tuple van floats 0-1 voor fitz."""
    hex_color = (hex_color or DEFAULT_TEXT_COLOR).lstrip("#")
    if len(hex_color) != 6:
        hex_color = DEFAULT_TEXT_COLOR.lstrip("#")
    return tuple(int(hex_color[i:i + 2], 16) / 255 for i in (0, 2, 4))


@dataclass
class TextAnnotation:
    page_num: int
    pdf_x: float
    pdf_y: float
    text: str
    font_size: int = DEFAULT_FONT_SIZE
    color: str = DEFAULT_TEXT_COLOR
    fontname: str = "helv"


@dataclass
class HighlightAnnotation:
    page_num: int
    quads: list = field(default_factory=list)  # fitz.Quad-objecten, PDF-coördinaten
    xref: int = 0  # PyMuPDF-xref van de live annotatie, voor verwijdering


@dataclass
class SignatureAnnotation:
    page_num: int
    pdf_x: float
    pdf_y: float
    width: float   # PDF-punten, ongezoomd
    height: float  # PDF-punten, ongezoomd
    image_bytes: bytes = b""  # PNG, klaar voor page.insert_image
