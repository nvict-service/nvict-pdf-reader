# -*- coding: utf-8 -*-
"""QApplication-opzet en entrypoint voor de PySide6-editie (fase 1)."""

import os
import sys

from PySide6.QtCore import QTimer, QTranslator
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from . import settings, theme, update_checker
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


def main(argv=None):
    app = QApplication(argv if argv is not None else sys.argv)
    app.setApplicationName("NVict Reader")
    app.setOrganizationName("NVict Service")

    _translator = _install_dutch_translator(app)  # referentie levend houden

    favicon_path = get_resource_path("favicon.ico")
    if os.path.exists(favicon_path):
        app.setWindowIcon(QIcon(favicon_path))

    theme.apply_theme(app, settings.get_theme_mode())

    window = MainWindow()

    # Direct een PDF openen als het als argument is meegegeven
    # (bv. via "Openen met" of drag-and-drop op de exe, in latere fases).
    args = argv if argv is not None else sys.argv
    for arg in args[1:]:
        if arg.lower().endswith(".pdf"):
            window.open_file(arg)
            break

    window.show()
    # Stille controle, pas na het opstarten - net als de tkinter-versie
    # (2s vertraging) zodat dit het laden van de PDF/UI niet vertraagt.
    QTimer.singleShot(2000, lambda: update_checker.check_for_updates(window, silent=True))
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
