# -*- coding: utf-8 -*-
"""Welkomscherm getoond zolang er geen tabblad open is.

Poort van NVict_Reader.py's welcome_frame/welcome_label (regel 1017:
exacte tekst "Welkom bij NVict Reader\\n\\nKlik op 'Openen' of druk op
Ctrl+O om een PDF te laden.") plus de recente-bestanden-sectie
(create_welcome_recent_section, regel 1727-1811): rijen met bestandsnaam
(links, accentkleur) en mapnaam (rechts, secundaire tekstkleur) op een
kaart die oplicht bij hover.

Gebruikt expliciete thema-kleuren (niet `palette(...)`) omdat de
QPalette-rollen "Mid"/"Link" niet door theme.py worden gezet en daardoor
in de praktijk te weinig contrast gaven (testfeedback: nauwelijks
leesbaar in zowel licht als donker thema).
"""

import os

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QToolButton, QVBoxLayout, QWidget

from . import settings, theme
from .resources import get_resource_path
from .i18n import tr

MAX_RECENT_SHOWN = 5
MAX_DIR_LENGTH = 55
LOGO_HEIGHT = 56


class _RecentFileRow(QFrame):
    def __init__(self, path, colors, on_click, on_remove, parent=None):
        super().__init__(parent)
        self._bg_normal = colors["BG_SECONDARY"]
        self._bg_hover = colors["BG_PRIMARY"]
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setStyleSheet(f"background-color: {self._bg_normal}; border-radius: 4px;")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 6, 10, 6)

        directory = os.path.dirname(path)
        if len(directory) > MAX_DIR_LENGTH:
            directory = "..." + directory[-(MAX_DIR_LENGTH - 3):]

        name_label = QLabel(f"\U0001F4C4  {os.path.basename(path)}", self)
        name_label.setStyleSheet(f"background: transparent; color: {colors['ACCENT_COLOR']};")
        layout.addWidget(name_label)
        layout.addStretch()

        dir_label = QLabel(directory, self)
        dir_label.setStyleSheet(f"background: transparent; color: {colors['TEXT_SECONDARY']};")
        layout.addWidget(dir_label)

        # Eigen knop: vangt zijn klik zelf af, dus opent het bestand niet.
        remove_btn = QToolButton(self)
        remove_btn.setText("✕")
        remove_btn.setToolTip(tr("Uit de lijst verwijderen"))
        remove_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        remove_btn.setStyleSheet(
            f"QToolButton {{ background: transparent; border: none; color: {colors['TEXT_SECONDARY']}; padding: 0 2px; }}"
            f"QToolButton:hover {{ color: {colors['ERROR_COLOR']}; }}"
        )
        remove_btn.clicked.connect(lambda: on_remove(path))
        layout.addWidget(remove_btn)

        for widget in (self, name_label, dir_label):
            widget.setCursor(Qt.CursorShape.PointingHandCursor)
        self.mousePressEvent = lambda _event: on_click(path)
        name_label.mousePressEvent = lambda _event: on_click(path)
        dir_label.mousePressEvent = lambda _event: on_click(path)

    def enterEvent(self, event):
        self.setStyleSheet(f"background-color: {self._bg_hover}; border-radius: 4px;")
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.setStyleSheet(f"background-color: {self._bg_normal}; border-radius: 4px;")
        super().leaveEvent(event)


class WelcomeWidget(QWidget):
    file_chosen = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
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
        self.recent_container.setFixedWidth(560)
        self.recent_layout = QVBoxLayout(self.recent_container)
        self.recent_layout.setContentsMargins(0, 24, 0, 0)
        self.recent_layout.setSpacing(4)
        outer.addWidget(self.recent_container)

        self.refresh()

    def refresh(self):
        """Herbouw de recente-bestanden-lijst - aanroepen na het openen van
        een bestand of een thema-wissel."""
        colors = theme.resolve_theme(settings.get_theme_mode())
        self.title_label.setStyleSheet(f"font-size: 16px; color: {colors['TEXT_SECONDARY']};")

        while self.recent_layout.count():
            item = self.recent_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                # takeAt() haalt de widget alleen uit de layout - zonder
                # hide() blijft hij op zijn oude positie zichtbaar totdat
                # deleteLater() een event-loop-tick later daadwerkelijk
                # verwijdert, wat een overlappende "geest"-rij gaf.
                widget.hide()
                widget.deleteLater()

        recent = [path for path in settings.get_recent_files() if os.path.exists(path)][:MAX_RECENT_SHOWN]
        if not recent:
            return

        header_row = QWidget(self.recent_container)
        header_layout = QHBoxLayout(header_row)
        header_layout.setContentsMargins(0, 0, 0, 4)
        header = QLabel(tr("Recente bestanden"), header_row)
        header.setStyleSheet(f"font-weight: bold; color: {colors['TEXT_SECONDARY']};")
        header_layout.addWidget(header)
        header_layout.addStretch()
        clear_link = QLabel(f'<a href="#" style="color: {colors["ACCENT_COLOR"]};">{tr("Lijst wissen")}</a>', header_row)
        clear_link.setToolTip(tr("Alle recente bestanden uit de lijst verwijderen"))
        clear_link.linkActivated.connect(lambda _href: self._clear_recent())
        header_layout.addWidget(clear_link)
        self.recent_layout.addWidget(header_row)

        for path in recent:
            row = _RecentFileRow(path, colors, self.file_chosen.emit, self._remove_recent, self.recent_container)
            self.recent_layout.addWidget(row)

    def _remove_recent(self, path):
        settings.remove_recent_file(path)
        self.refresh()

    def _clear_recent(self):
        settings.clear_recent_files()
        self.refresh()
