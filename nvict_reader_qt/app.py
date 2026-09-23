# -*- coding: utf-8 -*-
"""QApplication-opzet en entrypoint voor de PySide6-editie (fase 1)."""

import os
import sys

from PySide6.QtCore import QTimer, QTranslator
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from . import settings, single_instance, theme, update_checker
from .main_window import MainWindow
from .resources import get_resource_path


def _install_dutch_translator(app):
    """Vertaal Qt's eigen standaardknoppen (Ja/Nee/Annuleren/...) naar het
    Nederlands. PySide6 levert deze vertaling al kant-en-klaar mee - dit
    dekt QMessageBox/QFileDialog/QColorDialog/QPageSetupDialog e.d. in de
    hele app, zonder elke aanroep apart aan te moeten passen."""
    translations_dir = os.path.join(os.path.dirname(__import__("PySide6").__file__), "translations")
    translator = QTranslator(app)
    if translator.load("qtbase_nl", translations_dir):
        app.installTranslator(translator)
    return translator


def _activate_window(window):
    """Haalt het venster naar de voorgrond, ook als het geminimaliseerd is."""
    if window.isMinimized():
        window.showNormal()
    window.raise_()
    window.activateWindow()


def main(argv=None):
    app = QApplication(argv if argv is not None else sys.argv)
    app.setApplicationName("NVict Reader")
    app.setOrganizationName("NVict Service")

    # Direct een PDF openen als het als argument is meegegeven
    # (bv. via "Openen met" of dubbelklikken op een .pdf in de Verkenner).
    args = argv if argv is not None else sys.argv
    file_arg = None
    for arg in args[1:]:
        if arg.lower().endswith(".pdf"):
            file_arg = arg
            break

    # Draait de app al? Stuur het bestand dan naar die instance (als nieuwe
    # tab) i.p.v. hier zelf een nieuw venster te openen.
    if file_arg and single_instance.try_send_to_running_instance(file_arg):
        return 0

    _translator = _install_dutch_translator(app)  # referentie levend houden

    favicon_path = get_resource_path("favicon.ico")
    if os.path.exists(favicon_path):
        app.setWindowIcon(QIcon(favicon_path))

    theme.apply_theme(app, settings.get_theme_mode())

    window = MainWindow()

    instance_server = single_instance.SingleInstanceServer(app)

    def _on_file_received(path):
        window.open_file(path)
        _activate_window(window)

    instance_server.file_received.connect(_on_file_received)
    instance_server.start()

    if file_arg:
        window.open_file(file_arg)

    window.show()
    # Stille controle, pas na het opstarten - net als de tkinter-versie
    # (2s vertraging) zodat dit het laden van de PDF/UI niet vertraagt.
    QTimer.singleShot(2000, lambda: update_checker.check_for_updates(window, silent=True))
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
