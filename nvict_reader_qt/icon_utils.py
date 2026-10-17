# -*- coding: utf-8 -*-
"""Iconen uit icons/SVG, getekend in de kleur van het thema.

De SVG's gebruiken "currentColor" voor hun lijnen; die vervangen we door
de tekstkleur van het licht of donker thema. Zo zijn ze scherp op elke
schermschaal en hoeven we in donker thema geen bitmap om te keren (zoals
de oude PNG-iconen nodig hadden).
"""

import os
import tempfile

from PySide6.QtCore import QByteArray, QPointF, QRectF, Qt
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtGui import QColor, QImage, QPainter, QPen, QPixmap, QTransform

_temp_image_cache = {}

ICON_COLOR_LIGHT_THEME = "#1c1c1c"
ICON_COLOR_DARK_THEME = "#f0f0f0"


def svg_icon_path(name: str) -> str:
    """Pad naar icons/SVG/<name>.svg."""
    from .resources import get_resource_path
    return get_resource_path(os.path.join("icons", "SVG", f"{name}.svg"))


def svg_pixmap(path: str, color: str, size: int) -> QPixmap:
    """Render een SVG-icoon vierkant op `size` pixels in `color`."""
    with open(path, encoding="utf-8") as handle:
        svg = handle.read().replace("currentColor", color)
    renderer = QSvgRenderer(QByteArray(svg.encode("utf-8")))
    image = QImage(size, size, QImage.Format.Format_ARGB32_Premultiplied)
    image.fill(Qt.GlobalColor.transparent)
    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    renderer.render(painter, QRectF(0, 0, size, size))
    painter.end()
    return QPixmap.fromImage(image)


def _themed_svg_png(path: str, dark: bool, suffix: str, rotate: int = 0) -> str:
    """SVG -> tijdelijk PNG-bestand in themakleur. QSS (`image:`) accepteert
    alleen een bestandspad, geen QPixmap; daarom één keer wegschrijven."""
    key = (path, dark, suffix)
    cached = _temp_image_cache.get(key)
    if cached and os.path.exists(cached):
        return cached
    pixmap = svg_pixmap(path, ICON_COLOR_DARK_THEME if dark else ICON_COLOR_LIGHT_THEME, 64)
    if rotate:
        pixmap = pixmap.transformed(QTransform().rotate(rotate))
    name = os.path.splitext(os.path.basename(path))[0]
    temp_path = os.path.join(
        tempfile.gettempdir(), f"nvict_reader_qt_{suffix}_{name}_{'dark' if dark else 'light'}.png"
    )
    pixmap.save(temp_path, "PNG")
    _temp_image_cache[key] = temp_path
    return temp_path


def get_scrollbar_arrow_path(source_path: str, dark: bool) -> str:
    """Pijltje voor scrollbalken/getalvelden/keuzelijsten (de chevrons van
    Vorige/Volgende pagina), als pad voor de QSS."""
    return _themed_svg_png(source_path, dark, "arrow")


def get_toolbar_extension_arrow_path(source_path: str, dark: bool) -> str:
    """">"-chevron voor de overloopknop van de werkbalk: het "Volgende
    pagina"-chevron (wijst omlaag) een kwartslag gedraaid."""
    return _themed_svg_png(source_path, dark, "toolbar_more", rotate=-90)


INDICATOR_IMAGE_SIZE = 64  # groot getekend, Qt schaalt het scherp terug (ook bij 125-200%)


def get_indicator_image_path(kind: str, color: str) -> str:
    """Pad naar een vinkje ("check") of stip ("dot") in `color` voor de
    QSS van QCheckBox/QRadioButton. Een aangevinkt vakje was eerst helemaal
    met de accentkleur gevuld - onoverzichtelijk volgens gebruikersfeedback;
    nu een blauw vinkje in een wit vakje. Via een tijdelijk bestand omdat
    QSS alleen een pad accepteert."""
    key = (kind, color)
    cached = _temp_image_cache.get(key)
    if cached and os.path.exists(cached):
        return cached
    size = INDICATOR_IMAGE_SIZE
    image = QImage(size, size, QImage.Format.Format_ARGB32)
    image.fill(Qt.GlobalColor.transparent)
    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    if kind == "check":
        pen = QPen(QColor(color), size * 0.16)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)
        painter.drawPolyline([
            QPointF(size * 0.18, size * 0.52),
            QPointF(size * 0.42, size * 0.76),
            QPointF(size * 0.84, size * 0.26),
        ])
    else:
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(color))
        painter.drawEllipse(QPointF(size / 2, size / 2), size * 0.3, size * 0.3)
    painter.end()
    temp_path = os.path.join(tempfile.gettempdir(), f"nvict_reader_qt_{kind}_{color.lstrip('#')}.png")
    image.save(temp_path, "PNG")
    _temp_image_cache[key] = temp_path
    return temp_path
