# -*- coding: utf-8 -*-
"""Bewerken > Beveiligen met wachtwoord: een openingswachtwoord en/of een
eigenaarswachtwoord dat afdrukken/kopiëren/wijzigen beperkt.

De gebruiker moet expliciet aanvinken dat een vergeten wachtwoord niet te
herstellen is (AES-256: er bestaat geen achterdeur, ook niet bij NVict
Service) voordat er iets opgeslagen kan worden.
"""

from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QFrame,
    QGroupBox,
    QLabel,
    QLineEdit,
    QMessageBox,
    QVBoxLayout,
    QWidget,
)

from . import settings, theme
from .i18n import tr

MIN_PASSWORD_LENGTH = 4


class _PasswordPair(QWidget):
    """Wachtwoord + herhaling, met een gedeelde 'tonen'-schakelaar."""

    def __init__(self, parent=None):
        super().__init__(parent)
        form = QFormLayout(self)
        form.setContentsMargins(22, 0, 0, 0)
        self.first = QLineEdit(self)
        self.second = QLineEdit(self)
        for edit in (self.first, self.second):
            edit.setEchoMode(QLineEdit.EchoMode.Password)
        form.addRow(tr("Wachtwoord:"), self.first)
        form.addRow(tr("Herhalen:"), self.second)

    def set_visible_text(self, visible):
        mode = QLineEdit.EchoMode.Normal if visible else QLineEdit.EchoMode.Password
        for edit in (self.first, self.second):
            edit.setEchoMode(mode)

    def password(self):
        return self.first.text()

    def matches(self):
        return self.first.text() == self.second.text()


class ProtectPdfDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("Beveiligen met wachtwoord"))
        self.setMinimumWidth(480)
        colors = theme.resolve_theme(settings.get_theme_mode())

        layout = QVBoxLayout(self)

        intro = QLabel(tr(
            "De beveiligde versie wordt als nieuw bestand opgeslagen; "
            "het origineel blijft ongewijzigd."
        ), self)
        intro.setWordWrap(True)
        layout.addWidget(intro)

        # ── Openen ──
        self.open_check = QCheckBox(tr("Wachtwoord vereist om de PDF te openen"), self)
        layout.addWidget(self.open_check)
        self.open_pair = _PasswordPair(self)
        layout.addWidget(self.open_pair)

        # ── Beperken ──
        self.owner_check = QCheckBox(tr("Wijzigen beperken (eigenaarswachtwoord)"), self)
        layout.addWidget(self.owner_check)
        self.owner_pair = _PasswordPair(self)
        layout.addWidget(self.owner_pair)

        # Direct onder de invoervelden, zodat duidelijk is waar het bij hoort.
        self.show_check = QCheckBox(tr("Wachtwoorden tonen"), self)
        self.show_check.toggled.connect(self.open_pair.set_visible_text)
        self.show_check.toggled.connect(self.owner_pair.set_visible_text)
        layout.addWidget(self.show_check)

        self.perm_box = QGroupBox(tr("Zonder eigenaarswachtwoord toegestaan:"), self)
        perm_layout = QVBoxLayout(self.perm_box)
        self.allow_print = QCheckBox(tr("Afdrukken"), self.perm_box)
        self.allow_print.setChecked(True)
        self.allow_copy = QCheckBox(tr("Tekst en afbeeldingen kopiëren"), self.perm_box)
        self.allow_annotate = QCheckBox(tr("Opmerkingen toevoegen en formulieren invullen"), self.perm_box)
        for box in (self.allow_print, self.allow_copy, self.allow_annotate):
            perm_layout.addWidget(box)
        restrict_note = QLabel(tr(
            "Wijzigen van de inhoud is altijd uitgeschakeld. Beperkingen worden "
            "gerespecteerd door gangbare PDF-programma's, maar zijn geen absolute bescherming."
        ), self.perm_box)
        restrict_note.setWordWrap(True)
        restrict_note.setStyleSheet(f"color: {colors['TEXT_SECONDARY']};")
        perm_layout.addWidget(restrict_note)
        layout.addWidget(self.perm_box)

        # ── Waarschuwing ──
        warning = QFrame(self)
        warning.setObjectName("protectWarning")
        warning.setStyleSheet(
            f"#protectWarning {{ border: 2px solid {colors['ERROR_COLOR']}; border-radius: 6px; }}"
        )
        warning_layout = QVBoxLayout(warning)
        warning_text = QLabel(tr(
            "<b>⚠ Let op: een vergeten wachtwoord is NIET terug te halen.</b><br>"
            "De PDF wordt versleuteld met AES-256. Bent u het wachtwoord kwijt, dan kan "
            "niemand – ook NVict Service niet – het bestand nog openen of ontgrendelen. "
            "Bewaar het wachtwoord op een veilige plek, bijvoorbeeld in een wachtwoordmanager."
        ), warning)
        warning_text.setWordWrap(True)
        warning_text.setStyleSheet("border: none;")
        warning_layout.addWidget(warning_text)
        self.understand_check = QCheckBox(
            tr("Ik begrijp dat een vergeten wachtwoord niet te herstellen is"), warning
        )
        self.understand_check.setStyleSheet("border: none;")
        warning_layout.addWidget(self.understand_check)
        layout.addWidget(warning)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Cancel, self)
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText(tr("Annuleren"))
        self.ok_button = buttons.addButton(tr("Beveiligen en opslaan..."), QDialogButtonBox.ButtonRole.AcceptRole)
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        for box in (self.open_check, self.owner_check, self.understand_check):
            box.toggled.connect(self._update_state)
        self._update_state()

    def _update_state(self):
        self.open_pair.setEnabled(self.open_check.isChecked())
        self.owner_pair.setEnabled(self.owner_check.isChecked())
        self.perm_box.setEnabled(self.owner_check.isChecked())
        self.ok_button.setEnabled(
            (self.open_check.isChecked() or self.owner_check.isChecked())
            and self.understand_check.isChecked()
        )

    def _check_pair(self, pair, label):
        if len(pair.password()) < MIN_PASSWORD_LENGTH:
            return tr("Het {label} moet minstens {n} tekens lang zijn.", label=label, n=MIN_PASSWORD_LENGTH)
        if not pair.matches():
            return tr("De twee invoervelden van het {label} zijn niet gelijk.", label=label)
        return None

    def _on_accept(self):
        problems = []
        if self.open_check.isChecked():
            problems.append(self._check_pair(self.open_pair, tr("openingswachtwoord")))
        if self.owner_check.isChecked():
            problems.append(self._check_pair(self.owner_pair, tr("eigenaarswachtwoord")))
        if (self.open_check.isChecked() and self.owner_check.isChecked()
                and self.open_pair.password() == self.owner_pair.password()):
            problems.append(tr(
                "Het openingswachtwoord en het eigenaarswachtwoord moeten verschillen; "
                "anders heeft iedereen die de PDF kan openen ook alle rechten."
            ))
        problems = [p for p in problems if p]
        if problems:
            QMessageBox.warning(self, tr("Beveiligen met wachtwoord"), "\n\n".join(problems))
            return
        self.accept()

    def get_options(self) -> dict:
        owner = self.owner_check.isChecked()
        return {
            "open_password": self.open_pair.password() if self.open_check.isChecked() else "",
            "owner_password": self.owner_pair.password() if owner else "",
            "allow_print": self.allow_print.isChecked(),
            "allow_copy": self.allow_copy.isChecked(),
            "allow_annotate": self.allow_annotate.isChecked(),
        }
