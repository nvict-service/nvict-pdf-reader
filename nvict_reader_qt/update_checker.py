# -*- coding: utf-8 -*-
"""Automatische update-check, geport van NVict_Reader.py:5964-6266.

Bewust vereenvoudigd t.o.v. de tkinter-versie (gebruikerswens): één
duidelijke keuze ("Nu bijwerken" / "Later") in plaats van drie knoppen, en
geen melding dat het programma handmatig afgesloten moet worden - de
installer wordt gestart en het programma sluit daarna gewoon vanzelf
(`_finish_update` hieronder), zodat bijwerken zo min mogelijk uitleg of
tussenstappen vraagt.

Alle netwerk-I/O draait op een achtergrond-QThread; alleen de Qt-signalen
(`finished_ok`/`failed`/`progress`) raken de hoofdthread aan, net als het
bestaande print-voortgangspatroon in main_window.py.
"""

import hashlib
import json
import os
import tempfile
import urllib.error
import urllib.request

from PySide6.QtCore import QThread, QTimer, Signal
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QLabel,
    QMessageBox,
    QProgressDialog,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
)

from . import security
from .platform_win import is_packaged

APP_VERSION = "3.0"
UPDATE_CHECK_URL = "https://www.nvict.nl/software/updates/nvict_reader_version.json"


def _version_parts(text):
    parts = []
    for chunk in str(text).split("."):
        digits = ""
        for ch in chunk:
            if not ch.isdigit():
                break
            digits += ch
        parts.append(int(digits) if digits else 0)
    return parts


def _is_newer(latest, current):
    latest_parts = _version_parts(latest)
    current_parts = _version_parts(current)
    length = max(len(latest_parts), len(current_parts))
    latest_parts += [0] * (length - len(latest_parts))
    current_parts += [0] * (length - len(current_parts))
    return latest_parts > current_parts


class _CheckThread(QThread):
    finished_ok = Signal(dict)
    failed = Signal(str)

    def run(self):
        try:
            with urllib.request.urlopen(UPDATE_CHECK_URL, timeout=5) as response:
                data = json.loads(response.read().decode("utf-8"))
            self.finished_ok.emit(data)
        except Exception as exc:
            self.failed.emit(str(exc))


class _DownloadThread(QThread):
    progress = Signal(int, int)  # ontvangen bytes, totaal (0 = onbekend)
    finished_ok = Signal(str)
    failed = Signal(str)

    def __init__(self, download_url, version, expected_sha256, parent=None):
        super().__init__(parent)
        self.download_url = download_url
        self.version = version
        self.expected_sha256 = expected_sha256

    def run(self):
        try:
            # Eigen, verse map: een vast pad in %TEMP% is voorspelbaar en zou
            # door een ander proces vooraf klaargezet kunnen zijn.
            temp_dir = tempfile.mkdtemp(prefix="NVictReader_update_")
            filename = f"NVict_Reader_v{self.version}_Setup.exe"
            filepath = os.path.join(temp_dir, filename)

            def reporthook(block_num, block_size, total_size):
                self.progress.emit(block_num * block_size, max(total_size, 0))

            urllib.request.urlretrieve(self.download_url, filepath, reporthook)

            if self.expected_sha256:
                digest = hashlib.sha256()
                with open(filepath, "rb") as fh:
                    for chunk in iter(lambda: fh.read(1024 * 1024), b""):
                        digest.update(chunk)
                if digest.hexdigest().lower() != str(self.expected_sha256).strip().lower():
                    try:
                        os.remove(filepath)
                    except OSError:
                        pass
                    raise ValueError(
                        "De controlesom van het gedownloade bestand klopt niet. "
                        "De download is verwijderd en niet gestart."
                    )

            self.finished_ok.emit(filepath)
        except Exception as exc:
            self.failed.emit(str(exc))


class _UpdateAvailableDialog(QDialog):
    def __init__(self, parent, new_version, release_notes):
        super().__init__(parent)
        self.setWindowTitle("Update beschikbaar")
        self.setMinimumWidth(420)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            f"Er is een nieuwe versie van NVict Reader beschikbaar: <b>{new_version}</b> "
            f"(u gebruikt nu {APP_VERSION})."
        ))

        if release_notes:
            notes = QTextEdit(self)
            notes.setReadOnly(True)
            notes.setPlainText(release_notes)
            notes.setFixedHeight(140)
            layout.addWidget(notes)

        buttons = QDialogButtonBox(self)
        self.update_button = QPushButton("Nu bijwerken", self)
        self.update_button.setDefault(True)
        later_button = QPushButton("Later", self)
        buttons.addButton(self.update_button, QDialogButtonBox.ButtonRole.AcceptRole)
        buttons.addButton(later_button, QDialogButtonBox.ButtonRole.RejectRole)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)


def check_for_updates(parent, silent=False):
    """Controleer op een nieuwere versie. `silent=True`: geen melding als
    het al de nieuwste versie is of de server niet bereikbaar is (voor de
    stille controle bij het opstarten) - alleen tonen als er echt iets is.

    In de Microsoft Store-versie doet de Store de updates (en staat de
    Store-richtlijn een eigen updater niet toe), dus dan wordt er niets
    gecontroleerd of gedownload."""
    if is_packaged():
        if not silent:
            QMessageBox.information(
                parent, "Updates",
                "Deze versie van NVict Reader wordt automatisch bijgewerkt "
                "via de Microsoft Store.",
            )
        return
    thread = _CheckThread(parent)

    def on_ok(data):
        latest_version = str(data.get("version", "0.0"))
        download_url = data.get("download_url", "")
        release_notes = data.get("release_notes", "")
        expected_sha256 = data.get("sha256", "")

        if download_url and not security.is_trusted_update_url(download_url):
            if not silent:
                QMessageBox.critical(
                    parent, "Update geweigerd",
                    "De update-informatie verwijst naar een adres buiten "
                    "www.nvict.nl en is daarom genegeerd.\n\nDownload de update via de website.",
                )
            return

        if _is_newer(latest_version, APP_VERSION):
            _show_update_dialog(parent, latest_version, download_url, release_notes, expected_sha256)
        elif not silent:
            QMessageBox.information(parent, "Geen updates", f"U gebruikt al de nieuwste versie ({APP_VERSION}).")

    def on_fail(_error):
        if not silent:
            QMessageBox.critical(
                parent, "Verbindingsfout",
                "Kan niet verbinden met de update-server.\n\n"
                "Controleer uw internetverbinding en probeer het later opnieuw.",
            )

    thread.finished_ok.connect(on_ok)
    thread.failed.connect(on_fail)
    # Referentie levend houden zolang de thread loopt (anders GC'd Python
    # het object voordat de achtergrondthread klaar is).
    parent._update_check_thread = thread
    thread.start()


def _show_update_dialog(parent, new_version, download_url, release_notes, expected_sha256):
    dialog = _UpdateAvailableDialog(parent, new_version, release_notes)
    if dialog.exec() and download_url:
        _download_and_install(parent, download_url, new_version, expected_sha256)


def _download_and_install(parent, download_url, version, expected_sha256):
    if not security.is_trusted_update_url(download_url):
        QMessageBox.critical(
            parent, "Update geweigerd",
            "De update kan niet worden gedownload omdat het adres niet van "
            "www.nvict.nl komt.\n\nDownload de update via de website.",
        )
        return

    progress = QProgressDialog("Update downloaden...", "", 0, 0, parent)
    progress.setWindowTitle("Bijwerken")
    progress.setCancelButton(None)
    progress.setMinimumDuration(0)
    progress.show()

    thread = _DownloadThread(download_url, version, expected_sha256, parent)

    def on_progress(received, total):
        if total > 0:
            progress.setMaximum(total)
            progress.setValue(min(received, total))
        progress.setLabelText(f"Bezig met downloaden... ({received // 1024} KB)")

    def on_ok(filepath):
        progress.close()
        _finish_update(parent, filepath)

    def on_fail(error_msg):
        progress.close()
        QMessageBox.critical(
            parent, "Download mislukt",
            f"Kan de update niet downloaden:\n{error_msg}\n\n"
            "Probeer het later opnieuw of download de update handmatig via de website.",
        )

    thread.progress.connect(on_progress)
    thread.finished_ok.connect(on_ok)
    thread.failed.connect(on_fail)
    parent._update_download_thread = thread
    thread.start()


def _finish_update(parent, filepath):
    """Start de installer en sluit NVict Reader daarna vanzelf af - bewust
    geen "sluit het programma handmatig af"-melding (gebruikerswens): de
    gebruiker heeft met "Nu bijwerken" al bevestigd, een tweede bevestiging
    hier voegt niets toe."""
    if not os.path.exists(filepath):
        QMessageBox.critical(parent, "Fout", "Het gedownloade bestand is niet gevonden.")
        return
    try:
        os.startfile(filepath)
    except OSError as exc:
        QMessageBox.critical(parent, "Fout", f"Kan de installer niet starten:\n{exc}")
        return
    # Korte vertraging zodat Windows de installer daadwerkelijk gestart
    # heeft vóór dit programma zichzelf afsluit (parent.close() respecteert
    # de bestaande onopgeslagen-wijzigingen-check in closeEvent).
    QTimer.singleShot(800, parent.close)
