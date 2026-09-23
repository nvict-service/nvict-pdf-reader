# -*- coding: utf-8 -*-
"""Icoon-inversie voor donker thema.

Poort van tkinter's RGB-invert-met-alpha-behoud-truc (NVict_Reader.py:
894-901, _start_icon_load_thread): de icoonbestanden zelf zijn getekend
voor een licht thema (donkere lijnen op transparante achtergrond) en zijn
dus niet donker-thema-neutraal. In tkinter werd dit patroon gedupliceerd
tussen de hoofdtoolbar en de tekst-annotatie-editor (regel 4628-4633) -
hier één centrale functie voor de hele Qt-app.
"""

import os
import tempfile

from PySide6.QtGui import QImage, QPixmap


def invert_icon_colors(pixmap: QPixmap) -> QPixmap:
    """Geef een kopie van `pixmap` terug met geïnverteerde RGB-kanalen.

    Het alpha-kanaal blijft ongewijzigd, zodat de vorm van het icoon
    (transparante achtergrond) intact blijft - alleen de lijnkleur
    verandert van donker naar licht.
    """
    image = pixmap.toImage().convertToFormat(QImage.Format.Format_ARGB32)
    image.invertPixels(QImage.InvertMode.InvertRgb)
    return QPixmap.fromImage(image)


_scrollbar_arrow_cache = {}


def get_scrollbar_arrow_path(source_path: str, dark: bool) -> str:
    """Geef een bestandspad terug dat QSS's `image:`-property kan gebruiken
    voor de scrollbar-pijltjes (zelfde chevron-icoontjes als Vorige/
    Volgende pagina).

    QSS staat geen in-memory QPixmap/QIcon toe voor `image:` - alleen een
    pad - dus in donker thema wordt de (voor licht thema getekende, zwarte
    lijnen op transparant) bronafbeelding één keer geïnverteerd en als
    tijdelijk bestand gecachet; in licht thema is de originele afbeelding
    al bruikbaar.
    """
    if not dark:
        return source_path
    cached = _scrollbar_arrow_cache.get(source_path)
    if cached and os.path.exists(cached):
        return cached
    pixmap = invert_icon_colors(QPixmap(source_path))
    temp_path = os.path.join(tempfile.gettempdir(), f"nvict_reader_qt_dark_{os.path.basename(source_path)}")
    pixmap.save(temp_path, "PNG")
    _scrollbar_arrow_cache[source_path] = temp_path
    return temp_path
