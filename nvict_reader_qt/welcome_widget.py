# -*- coding: utf-8 -*-
"""Welkomscherm getoond zolang er geen tabblad open is.

Poort van NVict_Reader.py's welcome_frame/welcome_label (regel 1017:
exacte tekst "Welkom bij NVict Reader\\n\\nKlik op 'Openen' of druk op
Ctrl+O om een PDF te laden."), met daaronder de recente bestanden als
raster van thumbnails (eerste pagina), zoals in de Android-versie. Per
kaart een ✕ om hem uit de lijst te halen, plus "Lijst wissen" voor alles;
bij hover verschijnt het volledige pad (tooltip én regel onder het raster).

Thumbnails worden na het tonen van het scherm één voor één gerenderd
(eerste opbouw blijft direct) en op schijf gecachet per pad + wijzigtijd,
zodat een volgende start ze meteen heeft.

Gebruikt expliciete thema-kleuren (niet `palette(...)`) omdat de
QPalette-rollen "Mid"/"Link" niet door theme.py worden gezet en daardoor
in de praktijk te weinig contrast gaven (testfeedback: nauwelijks
leesbaar in zowel licht als donker thema).
"""

import hashlib
import os

from PySide6.QtCore import QStandardPaths, Qt, QTimer, Signal
from PySide6.QtGui import QFontMetrics, QPixmap
from PySide6.QtWidgets import QFrame, QGridLayout, QHBoxLayout, QLabel, QToolButton, QVBoxLayout, QWidget

from . import page_pixmap, settings, theme
from .document import get_fitz
from .resources import get_resource_path
from .i18n import tr

LOGO_HEIGHT = 56
THUMB_WIDTH = 120
THUMB_HEIGHT = 160
CARD_WIDTH = THUMB_WIDTH + 24
GRID_SPACING = 14
MAX_COLUMNS = 5


def _thumbnail_cache_dir() -> str:
    base = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.CacheLocation)
    return os.path.join(base, "recent_thumbnails")


def _thumbnail_cache_path(path, device_pixel_ratio) -> str:
    try:
        stat = os.stat(path)
        stamp = f"{stat.st_mtime_ns}:{stat.st_size}"
    except OSError:
        stamp = ""
    key = f"{os.path.normcase(os.path.abspath(path))}|{stamp}|{device_pixel_ratio:g}"
    return os.path.join(_thumbnail_cache_dir(), hashlib.sha1(key.encode("utf-8")).hexdigest() + ".png")


def _load_thumbnail(path, device_pixel_ratio):
    """Eerste pagina als pixmap, of None (versleuteld/beschadigd)."""
    cache_path = _thumbnail_cache_path(path, device_pixel_ratio)
    if os.path.exists(cache_path):
        pixmap = QPixmap(cache_path)
        if not pixmap.isNull():
            pixmap.setDevicePixelRatio(device_pixel_ratio)
            return pixmap

    try:
        doc = get_fitz().open(path)
    except Exception:
        return None
    try:
        if doc.needs_pass or len(doc) == 0:
            return None
        page = doc[0]
        bound = page.bound()
        if not bound.width or not bound.height:
            return None
        zoom = min(THUMB_WIDTH / bound.width, THUMB_HEIGHT / bound.height)
        pixmap = page_pixmap.render_page(page, zoom, device_pixel_ratio, reading=False)
    except Exception:
        return None
    finally:
        doc.close()

    try:
        os.makedirs(os.path.dirname(cache_path), exist_ok=True)
        pixmap.save(cache_path, "PNG")
    except Exception:
        pass  # cache is een bonus, geen vereiste
    return pixmap


def _prune_thumbnail_cache(keep_paths):
    cache_dir = _thumbnail_cache_dir()
    if not os.path.isdir(cache_dir):
        return
    keep = {os.path.basename(p) for p in keep_paths}
    for name in os.listdir(cache_dir):
        if name.endswith(".png") and name not in keep:
            try:
                os.remove(os.path.join(cache_dir, name))
            except OSError:
                pass


class _RecentFileCard(QFrame):
    def __init__(self, path, colors, on_click, on_remove, on_hover, parent=None):
        super().__init__(parent)
        self.path = path
        self._colors = colors
        self._on_click = on_click
        self._on_hover = on_hover
        self.setFixedWidth(CARD_WIDTH)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setToolTip(os.path.normpath(path))
        self.setObjectName("recentCard")
        self._apply_style(hover=False)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 8)
        layout.setSpacing(6)

        self.thumb_label = QLabel(self)
        self.thumb_label.setFixedSize(THUMB_WIDTH, THUMB_HEIGHT)
        self.thumb_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.thumb_label.setStyleSheet(f"background: transparent; color: {colors['TEXT_SECONDARY']}; font-size: 40px;")
        self.thumb_label.setText("\U0001F4C4")
        layout.addWidget(self.thumb_label, alignment=Qt.AlignmentFlag.AlignHCenter)

        name = os.path.basename(path)
        name_label = QLabel(self)
        name_label.setText(QFontMetrics(name_label.font()).elidedText(name, Qt.TextElideMode.ElideMiddle, THUMB_WIDTH))
        name_label.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        name_label.setStyleSheet(f"background: transparent; color: {colors['ACCENT_COLOR']};")
        layout.addWidget(name_label)

        # Eigen knop rechtsboven over de thumbnail: vangt zijn klik zelf af,
        # dus opent het bestand niet.
        self.remove_btn = QToolButton(self)
        self.remove_btn.setText("✕")
        self.remove_btn.setToolTip(tr("Uit de lijst verwijderen"))
        self.remove_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.remove_btn.setFixedSize(22, 22)
        self.remove_btn.setStyleSheet(
            f"QToolButton {{ background: {colors['BG_SECONDARY']}; border: 1px solid {colors['TEXT_SECONDARY']};"
            f" border-radius: 11px; color: {colors['TEXT_SECONDARY']}; padding: 0; }}"
            f"QToolButton:hover {{ color: {colors['ERROR_COLOR']}; border-color: {colors['ERROR_COLOR']}; }}"
        )
        self.remove_btn.clicked.connect(lambda: on_remove(path))
        self.remove_btn.move(CARD_WIDTH - 22 - 4, 4)
        self.remove_btn.raise_()

    def set_thumbnail(self, pixmap):
        if pixmap is None:
            return
        self.thumb_label.setText("")
        self.thumb_label.setStyleSheet(f"background: transparent; border: 1px solid {self._colors['BG_PRIMARY']};")
        self.thumb_label.setPixmap(pixmap)

    def _apply_style(self, hover):
        bg = self._colors["BG_PRIMARY"] if hover else self._colors["BG_SECONDARY"]
        border = self._colors["ACCENT_COLOR"] if hover else self._colors["BG_SECONDARY"]
        self.setStyleSheet(f"#recentCard {{ background-color: {bg}; border: 1px solid {border}; border-radius: 6px; }}")

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._on_click(self.path)
        super().mousePressEvent(event)

    def enterEvent(self, event):
        self._apply_style(hover=True)
        self._on_hover(self.path)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._apply_style(hover=False)
        self._on_hover(None)
        super().leaveEvent(event)


class WelcomeWidget(QWidget):
    file_chosen = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._cards = []
        self._thumbnail_memory = {}  # (pad, dpr) -> pixmap, overleeft refresh()
        self._pending_thumbnails = []
        self._columns = MAX_COLUMNS

        self._thumb_timer = QTimer(self)
        self._thumb_timer.setInterval(0)
        self._thumb_timer.timeout.connect(self._load_next_thumbnail)

        outer = QVBoxLayout(self)
        outer.setAlignment(Qt.AlignmentFlag.AlignCenter)

        logo_path = get_resource_path("logo.png")
        if os.path.exists(logo_path):
            logo_label = QLabel(self)
            logo_pixmap = QPixmap(logo_path).scaledToHeight(
                LOGO_HEIGHT, Qt.TransformationMode.SmoothTransformation
            )
            logo_label.setPixmap(logo_pixmap)
            logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            logo_label.setStyleSheet("margin-bottom: 12px;")
            outer.addWidget(logo_label)

        self.title_label = QLabel(
            tr("Welkom bij NVict Reader\n\nKlik op 'Openen' of druk op Ctrl+O om een PDF te laden."), self
        )
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label.setStyleSheet("font-size: 16px;")
        outer.addWidget(self.title_label)

        self.recent_container = QWidget(self)
        recent_layout = QVBoxLayout(self.recent_container)
        recent_layout.setContentsMargins(0, 24, 0, 0)
        recent_layout.setSpacing(8)

        self.header_row = QWidget(self.recent_container)
        header_layout = QHBoxLayout(self.header_row)
        header_layout.setContentsMargins(0, 0, 0, 0)
        self.header_label = QLabel(tr("Recente bestanden"), self.header_row)
        header_layout.addWidget(self.header_label)
        header_layout.addStretch()
        self.clear_link = QLabel(self.header_row)
        self.clear_link.setToolTip(tr("Alle recente bestanden uit de lijst verwijderen"))
        self.clear_link.linkActivated.connect(lambda _href: self._clear_recent())
        header_layout.addWidget(self.clear_link)
        recent_layout.addWidget(self.header_row)

        self.grid_widget = QWidget(self.recent_container)
        self.grid = QGridLayout(self.grid_widget)
        self.grid.setContentsMargins(0, 0, 0, 0)
        self.grid.setHorizontalSpacing(GRID_SPACING)
        self.grid.setVerticalSpacing(GRID_SPACING)
        recent_layout.addWidget(self.grid_widget, alignment=Qt.AlignmentFlag.AlignLeft)

        # Vaste hoogte, ook als hij leeg is: anders verspringt het raster
        # bij elke hover.
        self.path_label = QLabel(self.recent_container)
        self.path_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.path_label.setFixedHeight(QFontMetrics(self.path_label.font()).height() + 4)
        recent_layout.addWidget(self.path_label)

        outer.addWidget(self.recent_container, alignment=Qt.AlignmentFlag.AlignHCenter)

        self.refresh()

    def refresh(self):
        """Herbouw de recente-bestanden-kaarten - aanroepen na het openen
        van een bestand of een thema-wissel."""
        colors = theme.resolve_theme(settings.get_theme_mode())
        self.title_label.setStyleSheet(f"font-size: 16px; color: {colors['TEXT_SECONDARY']};")
        self.header_label.setStyleSheet(f"font-weight: bold; color: {colors['TEXT_SECONDARY']};")
        self.clear_link.setText(f'<a href="#" style="color: {colors["ACCENT_COLOR"]};">{tr("Lijst wissen")}</a>')
        self.path_label.setStyleSheet(f"color: {colors['TEXT_SECONDARY']};")
        self.path_label.clear()

        self._thumb_timer.stop()
        for card in self._cards:
            self.grid.removeWidget(card)
            # removeWidget() haalt de kaart alleen uit de layout - zonder
            # hide() blijft hij op zijn oude positie zichtbaar totdat
            # deleteLater() een event-loop-tick later daadwerkelijk
            # verwijdert, wat een overlappende "geest"-kaart gaf.
            card.hide()
            card.deleteLater()
        self._cards = []

        recent = [path for path in settings.get_recent_files() if os.path.exists(path)]
        self.recent_container.setVisible(bool(recent))

        dpr = self.devicePixelRatioF()
        for path in recent:
            card = _RecentFileCard(
                path, colors, self.file_chosen.emit, self._remove_recent, self._show_path, self.grid_widget
            )
            pixmap = self._thumbnail_memory.get((path, dpr))
            if pixmap is not None:
                card.set_thumbnail(pixmap)
            self._cards.append(card)
        self._columns = self._fitting_columns()
        self._layout_cards()

        self._pending_thumbnails = [card for card in self._cards if (card.path, dpr) not in self._thumbnail_memory]
        if self._pending_thumbnails:
            self._thumb_timer.start()
        else:
            self._prune_cache(recent, dpr)

    def _fitting_columns(self):
        available = max(CARD_WIDTH, self.width() - 40)
        return max(1, min(MAX_COLUMNS, (available + GRID_SPACING) // (CARD_WIDTH + GRID_SPACING)))

    def _layout_cards(self):
        for index, card in enumerate(self._cards):
            self.grid.addWidget(card, index // self._columns, index % self._columns)
            card.show()
        path_width = self._columns * CARD_WIDTH + (self._columns - 1) * GRID_SPACING
        self.path_label.setFixedWidth(path_width)

    def _load_next_thumbnail(self):
        if not self._pending_thumbnails:
            self._thumb_timer.stop()
            self._prune_cache([card.path for card in self._cards], self.devicePixelRatioF())
            return
        card = self._pending_thumbnails.pop(0)
        dpr = self.devicePixelRatioF()
        pixmap = _load_thumbnail(card.path, dpr)
        self._thumbnail_memory[(card.path, dpr)] = pixmap
        card.set_thumbnail(pixmap)

    def _prune_cache(self, paths, dpr):
        _prune_thumbnail_cache([_thumbnail_cache_path(path, dpr) for path in paths])
        self._thumbnail_memory = {key: value for key, value in self._thumbnail_memory.items() if key[0] in paths}

    def _show_path(self, path):
        if path is None:
            self.path_label.clear()
            return
        metrics = QFontMetrics(self.path_label.font())
        self.path_label.setText(metrics.elidedText(os.path.normpath(path), Qt.TextElideMode.ElideMiddle, self.path_label.width()))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        columns = self._fitting_columns()
        if columns != self._columns and self._cards:
            self._columns = columns
            for card in self._cards:
                self.grid.removeWidget(card)
            self._layout_cards()

    def _remove_recent(self, path):
        settings.remove_recent_file(path)
        self.refresh()

    def _clear_recent(self):
        settings.clear_recent_files()
        self.refresh()
