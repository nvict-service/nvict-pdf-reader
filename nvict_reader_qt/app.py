# -*- coding: utf-8 -*-
"""QApplication-opzet en entrypoint voor de PySide6-editie (fase 1)."""

import os
import sys

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from . import settings, theme
from .main_window import MainWindow
from .resources import get_resource_path


def main(argv=None):
    app = QApplication(argv if argv is not None else sys.argv)
    app.setApplicationName("NVict Reader")
    app.setOrganizationName("NVict Service")

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
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
