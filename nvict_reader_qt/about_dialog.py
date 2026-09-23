# -*- coding: utf-8 -*-
""""Over NVict Reader"-dialoog, geport van NVict_Reader.py:5831-5958.

Zelfde opbouw als de tkinter-versie (accentkleur-header, logo op een witte
achtergrond, functielijst, links naar nvict.nl en de Flaticon-attributie),
maar de functielijst is bijgewerkt naar wat de PySide6-editie nu
daadwerkelijk kan (undo, zoeken, boekweergave/volledig scherm,
handtekening) in plaats van de kortere tkinter-lijst uit 2026. Aanvullend
op verzoek: een disclaimer ("gratis, zonder garantie") en een link naar de
GitHub-broncode - dat had de tkinter-versie niet.
"""

import os
from datetime import datetime

from PySide6.QtCore import Qt
from PySide6.QtGui import QPainter, QPixmap
from PySide6.QtWidgets import QDialog, QFrame, QLabel, QPushButton, QVBoxLayout

from . import settings, theme
from .resources import get_resource_path
from .update_checker import APP_VERSION

FLATICON_URL = "https://www.flaticon.com/free-icons/page"
GITHUB_URL = "https://github.com/nvict-service/nvict-pdf-reader"
DISCLAIMER_TEXT = (
    "NVict Reader is gratis software, aangeboden zonder enige garantie. "
    "Gebruik is op eigen risico."
)

FEATURES = [
    "PDF's openen en bekijken",
    "Tekst selecteren en kopiëren",
    "Markeren en tekst toevoegen",
    "Formulieren invullen",
    "Handtekening plaatsen",
    "Pagina's exporteren en samenvoegen",
    "Pagina's roteren",
    "Afdrukken met opties",
    "Zoeken in documenten",
    "Volledig scherm en boekweergave",
]


class AboutDialog(QDialog):
    def __init__(self, parent, website_url: str):
        super().__init__(parent)
        self.setWindowTitle("Over NVict Reader")
        self.setFixedWidth(420)
        # Expliciete thema-kleuren i.p.v. palette(mid)/palette(link): die
        # QPalette-rollen worden door theme.py niet gezet en gaven daardoor
        # (zelfde probleem als eerder bij het welkomscherm) te weinig
        # contrast in zowel licht als donker thema.
        colors = theme.resolve_theme(settings.get_theme_mode())

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        header = QFrame(self)
        header.setStyleSheet(f"background-color: {colors['ACCENT_COLOR']};")
        header_layout = QVBoxLayout(header)
        header_label = QLabel("Over NVict Reader", header)
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_label.setStyleSheet("color: white; font-size: 15px; font-weight: bold; background: transparent;")
        header_layout.addWidget(header_label)
        header_layout.setContentsMargins(20, 15, 20, 15)
        outer.addWidget(header)

        content = QVBoxLayout()
        content.setContentsMargins(30, 20, 30, 20)
        content.setAlignment(Qt.AlignmentFlag.AlignHCenter)

        logo_path = get_resource_path("logo.png")
        if os.path.exists(logo_path):
            logo_label = QLabel(self)
            canvas = QPixmap(100, 100)
            canvas.fill(Qt.GlobalColor.white)
            logo_pixmap = QPixmap(logo_path).scaled(
                80, 80, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation
            )
            painter = QPainter(canvas)
            painter.drawPixmap((100 - logo_pixmap.width()) // 2, (100 - logo_pixmap.height()) // 2, logo_pixmap)
            painter.end()
            logo_label.setPixmap(canvas)
            logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            content.addWidget(logo_label)

        title_label = QLabel("NVict Reader", self)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("font-size: 16px; font-weight: bold; margin-top: 10px;")
        content.addWidget(title_label)

        version_label = QLabel(f"Versie {APP_VERSION}", self)
        version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        content.addWidget(version_label)

        copyright_label = QLabel(f"© {datetime.now().year} NVict Service", self)
        copyright_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        content.addWidget(copyright_label)

        disclaimer_label = QLabel(DISCLAIMER_TEXT, self)
        disclaimer_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        disclaimer_label.setWordWrap(True)
        disclaimer_label.setStyleSheet(
            f"color: {colors['TEXT_SECONDARY']}; font-size: 11px; margin: 6px 10px 15px 10px;"
        )
        content.addWidget(disclaimer_label)

        features_box = QFrame(self)
        features_box.setStyleSheet(f"background-color: {colors['BG_SECONDARY']}; border-radius: 6px;")
        features_layout = QVBoxLayout(features_box)
        features_title = QLabel("Functies:", features_box)
        features_title.setStyleSheet("font-weight: bold; background: transparent;")
        features_layout.addWidget(features_title)
        for feature in FEATURES:
            label = QLabel(f"✓ {feature}", features_box)
            label.setStyleSheet("background: transparent;")
            features_layout.addWidget(label)
        content.addWidget(features_box)

        link_label = QLabel(f'<a href="{website_url}">www.nvict.nl</a>', self)
        link_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        link_label.setStyleSheet("margin-top: 15px;")
        link_label.setOpenExternalLinks(True)
        content.addWidget(link_label)

        github_label = QLabel(f'<a href="{GITHUB_URL}">Broncode op GitHub</a>', self)
        github_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        github_label.setOpenExternalLinks(True)
        content.addWidget(github_label)

        credit_label = QLabel(f'<a href="{FLATICON_URL}">Iconen door Freepik - Flaticon</a>', self)
        credit_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        credit_label.setStyleSheet(f"color: {colors['TEXT_SECONDARY']}; font-size: 11px;")
        credit_label.setOpenExternalLinks(True)
        content.addWidget(credit_label)

        outer.addLayout(content)

        footer = QFrame(self)
        footer.setStyleSheet(f"background-color: {colors['BG_SECONDARY']};")
        footer_layout = QVBoxLayout(footer)
        close_button = QPushButton("Sluiten", footer)
        close_button.clicked.connect(self.accept)
        footer_layout.addWidget(close_button, alignment=Qt.AlignmentFlag.AlignCenter)
        outer.addWidget(footer)
