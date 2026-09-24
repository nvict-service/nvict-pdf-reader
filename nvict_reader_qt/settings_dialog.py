# -*- coding: utf-8 -*-
"""Instellingenvenster: thema-keuze + standaard-PDF-viewer.

Grotendeels nieuwe UI - tkinter had hier geen werkend equivalent
(toggle_theme was dode code zonder UI-koppeling, zie NVict_Reader.py:1581).
Het "instellen als standaard"-gedeelte hergebruikt platform_win.py, dat al
sinds fase 1 klaarstond maar nog nergens werd aangeroepen.
"""

from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QDialogButtonBox,
    QGroupBox,
    QLabel,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QDialog,
)

from . import settings
from .platform_win import DefaultPDFHandler, is_packaged

THEME_LABELS = ["Licht", "Donker", "Systeemstandaard"]


class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Instellingen")
        self.setMinimumWidth(360)

        layout = QVBoxLayout(self)

        theme_box = QGroupBox("Thema", self)
        theme_layout = QVBoxLayout(theme_box)
        self.theme_group = QButtonGroup(self)
        current_mode = settings.get_theme_mode()
        self.theme_buttons = {}
        for label in THEME_LABELS:
            rb = QRadioButton(label, theme_box)
            rb.setChecked(label == current_mode)
            self.theme_group.addButton(rb)
            self.theme_buttons[rb] = label
            theme_layout.addWidget(rb)
        layout.addWidget(theme_box)

        view_box = QGroupBox("Weergave", self)
        view_layout = QVBoxLayout(view_box)
        self.show_thumbnails_check = QCheckBox("Pagina's (miniaturen) tonen bij het openen van een document", view_box)
        self.show_thumbnails_check.setChecked(settings.get_show_thumbnails_default())
        view_layout.addWidget(self.show_thumbnails_check)
        layout.addWidget(view_box)

        default_box = QGroupBox("Standaard PDF-viewer", self)
        default_layout = QVBoxLayout(default_box)
        is_default = DefaultPDFHandler.is_default_pdf_handler()
        self.default_status_label = QLabel(
            "NVict Reader is momenteel de standaard PDF-viewer." if is_default
            else "NVict Reader is nog niet de standaard PDF-viewer.",
            default_box,
        )
        default_layout.addWidget(self.default_status_label)
        set_default_btn = QPushButton("Instellen als standaard...", default_box)
        set_default_btn.clicked.connect(self._set_as_default)
        default_layout.addWidget(set_default_btn)
        layout.addWidget(default_box)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _set_as_default(self):
        # De Store-versie krijgt de .pdf-koppeling uit het package-manifest;
        # registerschrijfacties worden daar gevirtualiseerd en hebben geen effect.
        if not is_packaged():
            DefaultPDFHandler.register_open_with()
        reply = QMessageBox.question(
            self, "Standaard app instellen",
            "Om NVict Reader als standaard in te stellen, opent Windows nu het instellingenmenu.\n\n"
            "1. Zoek '.pdf' in de lijst of klik op de huidige standaard app.\n"
            "2. Selecteer 'NVict Reader' in de lijst.\n"
            "3. Klik op 'Als standaard instellen'.\n\n"
            "Wilt u de instellingen nu openen?",
        )
        if reply == QMessageBox.StandardButton.Yes:
            DefaultPDFHandler.open_windows_default_apps_pdf()

    def get_theme_mode(self) -> str:
        for rb, label in self.theme_buttons.items():
            if rb.isChecked():
                return label
        return "Systeemstandaard"

    def get_show_thumbnails_default(self) -> bool:
        return self.show_thumbnails_check.isChecked()
