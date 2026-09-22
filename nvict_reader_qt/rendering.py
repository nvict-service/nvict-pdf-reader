# -*- coding: utf-8 -*-
"""Pagina-layout, zichtbaarheids-detectie en pixmap-cache.

Poort van de tkinter-versie (NVict_Reader.py regel 2138-2365:
_compute_page_layout / _visible_page_numbers / _render_visible_pages /
_trim_page_cache), losgemaakt van het canvas-object zodat deze logica
zonder Qt-widgets te instantiëren getest kan worden. Book-modus (twee
pagina's naast elkaar) is hier nog niet overgenomen - fase 1 kent alleen
doorlopend enkelkoloms-scrollen.
"""

from collections import namedtuple

RENDER_MARGIN_PAGES = 1   # extra pagina's boven en onder het beeld
RENDER_CACHE_PAGES = 12   # maximaal aantal bitmaps dat we vasthouden

X_MARGIN = 20
Y_START = 20
PAGE_SPACING = 20

PageLayoutEntry = namedtuple("PageLayoutEntry", ["page_num", "x", "y", "width", "height"])

WORD_CACHE_PAGES = 20  # aantal pagina's waarvan we de woordenlijst vasthouden

WordBox = namedtuple("WordBox", ["text", "x0", "y0", "x1", "y1"])


def compute_page_layout(pdf_document, zoom_level):
    """Bereken pagina-posities en totale documentgrootte voor doorlopend scrollen.

    Geeft (layout, total_width, total_height) terug. `layout` is een lijst
    van PageLayoutEntry, één per pagina, in documentvolgorde.
    """
    layout = []
    y = Y_START
    max_width = 0

    for page_num in range(len(pdf_document)):
        bound = pdf_document[page_num].bound()
        w = int(bound.width * zoom_level)
        h = int(bound.height * zoom_level)
        layout.append(PageLayoutEntry(page_num, X_MARGIN, y, w, h))
        max_width = max(max_width, X_MARGIN + w)
        y += h + PAGE_SPACING

    total_width = max_width + X_MARGIN
    total_height = y + Y_START
    return layout, total_width, total_height


def fit_width_zoom(pdf_document, viewport_width):
    """Bereken het zoomniveau waarbij pagina 0 net binnen `viewport_width` past."""
    usable_width = max(viewport_width - 2 * X_MARGIN, 100)
    page_width = pdf_document[0].bound().width
    if page_width <= 0:
        return 1.0
    return usable_width / page_width


def visible_page_numbers(layout, viewport_top, viewport_height, current_page=0, margin=RENDER_MARGIN_PAGES):
    """Geef de paginanummers terug die in beeld staan, plus een marge.

    `layout` is een lijst van PageLayoutEntry. `viewport_top`/`viewport_height`
    zijn scene-coördinaten (bv. via QGraphicsView.mapToScene(viewport rect)).
    """
    if not layout:
        return []

    bottom = viewport_top + viewport_height
    visible_idx = [
        i for i, entry in enumerate(layout)
        if entry.y <= bottom and (entry.y + entry.height) >= viewport_top
    ]

    if not visible_idx:
        # Buiten elk bereik (bv. net na een zoomwissel): val terug op de
        # pagina waar de gebruiker naartoe genavigeerd is.
        current = min(max(current_page, 0), len(layout) - 1)
        visible_idx = [i for i, entry in enumerate(layout) if entry.page_num == current] or [0]

    first = max(0, min(visible_idx) - margin)
    last = min(len(layout) - 1, max(visible_idx) + margin)
    return [layout[i].page_num for i in range(first, last + 1)]


def get_page_words(pdf_document, page_num):
    """Geef de woorden van een pagina terug in PDF-coördinaten (ongezoomd).

    Poort van het `page.get_text("words")`-gebruik in NVict_Reader.py
    (regel 2298-2306), los van canvas/zoom - de aanroeper rekent zelf om
    naar scene-coördinaten via de PageLayoutEntry van deze pagina.
    """
    page = pdf_document[page_num]
    return [WordBox(w[4], w[0], w[1], w[2], w[3]) for w in page.get_text("words")]


def trim_page_cache(cache: dict, keep, anchor, max_pages=RENDER_CACHE_PAGES):
    """Gooi bitmaps weg van pagina's die ver buiten beeld liggen.

    Zonder deze begrenzing groeit het geheugengebruik bij het doorbladeren
    van een lang document ongelimiteerd: elke gerenderde pagina is een
    volledige bitmap op zoomniveau. Muteert `cache` in-place.
    """
    if len(cache) <= max_pages:
        return
    keep_set = set(keep)
    for page_num in sorted(cache.keys(), key=lambda p: -abs(p - anchor)):
        if len(cache) <= max_pages:
            break
        if page_num not in keep_set:
            cache.pop(page_num, None)
