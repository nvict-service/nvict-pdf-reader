# -*- coding: utf-8 -*-
"""Help > Fout melden: de gebruiker beschrijft het probleem en we stellen
een e-mail op in het eigen mailprogramma (mailto:), met systeeminformatie
die in het venster zelf zichtbaar is - de gebruiker ziet dus precies wat
er verstuurd wordt en verstuurt zelf. Geen server of account nodig.

Zonder (gekoppeld) mailprogramma zet het de tekst op het klembord en
noemen we het adres, zodat het via webmail alsnog kan.
"""

import platform
from urllib.parse import quote

from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices, QGuiApplication
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QVBoxLayout,
)

from . import i18n, platform_win
from .i18n import tr
from .update_checker import APP_VERSION

REPORT_ADDRESS = "software@nvict.nl"


def _system_info(widget) -> str:
    screen = widget.screen() or QGuiApplication.primaryScreen()
    scale = f"{round(screen.devicePixelRatio() * 100)}%" if screen else "?"
    size = f"{screen.size().width()}x{screen.size().height()}" if screen else "?"
    install = "Microsoft Store" if platform_win.is_packaged() else tr("Installatieprogramma")
    windows = f"{platform.system()} {platform.release()} ({platform.version()})"
    return "\n".join([
        f"NVict Reader {APP_VERSION} ({install})",
        f"{windows}",
        tr("Scherm: {size}, schaal {scale}", size=size, scale=scale),
        tr("Taal: {language}", language=i18n.current()),
    ])


class BugReportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("Fout melden"))
        self.setMinimumWidth(460)

        layout = QVBoxLayout(self)
        intro = QLabel(tr(
            "Beschrijf wat er misging en wat u op dat moment deed. Met 'E-mail opstellen' "
            "opent uw mailprogramma een bericht aan {address}; u verstuurt het zelf.",
            address=REPORT_ADDRESS,
        ), self)
        intro.setWordWrap(True)
        layout.addWidget(intro)

        self.description_edit = QPlainTextEdit(self)
        self.description_edit.setPlaceholderText(tr("Bijvoorbeeld: bij het openen van een bepaalde PDF blijft het scherm leeg..."))
        self.description_edit.setMinimumHeight(140)
        self.description_edit.textChanged.connect(self._update_send_enabled)
        layout.addWidget(self.description_edit)

        self.include_info_check = QCheckBox(tr("Systeeminformatie meesturen (helpt bij het oplossen)"), self)
        self.include_info_check.setChecked(True)
        layout.addWidget(self.include_info_check)

        self.system_info = _system_info(self)
        info_label = QLabel(self.system_info, self)
        info_label.setStyleSheet("margin-left: 22px;")
        self.include_info_check.toggled.connect(info_label.setEnabled)
        layout.addWidget(info_label)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Cancel, self)
        self.send_button = buttons.addButton(tr("E-mail opstellen"), QDialogButtonBox.ButtonRole.AcceptRole)
        buttons.accepted.connect(self._send)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self._update_send_enabled()

    def _update_send_enabled(self):
        self.send_button.setEnabled(bool(self.description_edit.toPlainText().strip()))

    def _body(self) -> str:
        body = self.description_edit.toPlainText().strip()
        if self.include_info_check.isChecked():
            body += "\n\n---\n" + self.system_info
        return body

    def _send(self):
        subject = tr("Foutmelding NVict Reader {version}", version=APP_VERSION)
        body = self._body()
        url = QUrl(f"mailto:{REPORT_ADDRESS}?subject={quote(subject)}&body={quote(body)}")
        if QDesktopServices.openUrl(url):
            self.accept()
            return
        QGuiApplication.clipboard().setText(f"{subject}\n\n{body}")
        QMessageBox.information(self, tr("Fout melden"), tr(
            "Er is geen mailprogramma gevonden. De melding staat op het klembord; "
            "plak hem in een e-mail aan {address}.",
            address=REPORT_ADDRESS,
        ))
        self.accept()
