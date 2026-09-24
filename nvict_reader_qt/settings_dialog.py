# -*- coding: utf-8 -*-
"""Instellingenvenster: thema, taal, weergave + standaard-PDF-viewer.

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

from . import i18n, settings
from .i18n import tr
from .platform_win import DefaultPDFHandler, is_packaged

# Opgeslagen waarden (blijven Nederlands, ook in de Engelse versie - theme.py
# rekent ermee); alleen het label op het scherm wordt vertaald.
THEME_LABELS = ["Licht", "Donker", "Systeemstandaard"]

LANGUAGE_OPTIONS = [
    (i18n.LANGUAGE_SYSTEM, "Systeemstandaard (taal van Windows)"),
    (i18n.LANGUAGE_DUTCH, "Nederlands"),
    (i18n.LANGUAGE_ENGLISH, "English"),
]


class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("Instellingen"))
        self.setMinimumWidth(360)

        layout = QVBoxLayout(self)

        theme_box = QGroupBox(tr("Thema"), self)
        theme_layout = QVBoxLayout(theme_box)
        self.theme_group = QButtonGroup(self)
        current_mode = settings.get_theme_mode()
        self.theme_buttons = {}
        for label in THEME_LABELS:
            rb = QRadioButton(tr(label), theme_box)
            rb.setChecked(label == current_mode)
            self.theme_group.addButton(rb)
            self.theme_buttons[rb] = label
            theme_layout.addWidget(rb)
        layout.addWidget(theme_box)

        language_box = QGroupBox(tr("Taal"), self)
        language_layout = QVBoxLayout(language_box)
        self.language_group = QButtonGroup(self)
        self._initial_language = settings.get_language()
        self.language_buttons = {}
        for value, label in LANGUAGE_OPTIONS:
            # "Nederlands"/"English" staan altijd in hun eigen taal, zodat je
            # de taal terugvindt ook als je de huidige niet leest.
            rb = QRadioButton(tr(label) if value == i18n.LANGUAGE_SYSTEM else label, language_box)
            rb.setChecked(value == self._initial_language)
            self.language_group.addButton(rb)
            self.language_buttons[rb] = value
            language_layout.addWidget(rb)
        layout.addWidget(language_box)

        view_box = QGroupBox(tr("Weergave"), self)
        view_layout = QVBoxLayout(view_box)
        self.show_thumbnails_check = QCheckBox(tr("Pagina's (miniaturen) tonen bij het openen van een document"), view_box)
        self.show_thumbnails_check.setChecked(settings.get_show_thumbnails_default())
        view_layout.addWidget(self.show_thumbnails_check)
        self.clear_recent_btn = QPushButton(tr("Lijst met recente bestanden wissen"), view_box)
        self.clear_recent_btn.clicked.connect(self._clear_recent_files)
        self.clear_recent_btn.setEnabled(bool(settings.get_recent_files()))
        view_layout.addWidget(self.clear_recent_btn)
        layout.addWidget(view_box)

        default_box = QGroupBox(tr("Standaard PDF-viewer"), self)
        default_layout = QVBoxLayout(default_box)
        is_default = DefaultPDFHandler.is_default_pdf_handler()
        self.default_status_label = QLabel(
            tr("NVict Reader is momenteel de standaard PDF-viewer.") if is_default
            else tr("NVict Reader is nog niet de standaard PDF-viewer."),
            default_box,
        )
        default_layout.addWidget(self.default_status_label)
        set_default_btn = QPushButton(tr("Instellen als standaard..."), default_box)
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
            self, tr("Standaard app instellen"),
            tr("Om NVict Reader als standaard in te stellen, opent Windows nu het instellingenmenu.\n\n"
               "1. Zoek '.pdf' in de lijst of klik op de huidige standaard app.\n"
               "2. Selecteer 'NVict Reader' in de lijst.\n"
               "3. Klik op 'Als standaard instellen'.\n\n"
               "Wilt u de instellingen nu openen?"),
        )
        if reply == QMessageBox.StandardButton.Yes:
            DefaultPDFHandler.open_windows_default_apps_pdf()

    def _clear_recent_files(self):
        # Direct uitgevoerd (niet pas bij OK): het is een losse actie, geen instelling.
        settings.clear_recent_files()
        self.clear_recent_btn.setEnabled(False)
        self.clear_recent_btn.setText(tr("Lijst met recente bestanden is gewist"))

    def get_theme_mode(self) -> str:
        for rb, label in self.theme_buttons.items():
            if rb.isChecked():
                return label
        return "Systeemstandaard"

    def get_show_thumbnails_default(self) -> bool:
        return self.show_thumbnails_check.isChecked()

    def get_language(self) -> str:
        for rb, value in self.language_buttons.items():
            if rb.isChecked():
                return value
        return i18n.LANGUAGE_SYSTEM

    def language_changed(self) -> bool:
        """True als de gekozen taal na een herstart anders uitpakt dan nu."""
        return i18n.resolve(self.get_language()) != i18n.current()
