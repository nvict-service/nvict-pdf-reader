# -*- coding: utf-8 -*-
"""QApplication-opzet en entrypoint voor de PySide6-editie (fase 1)."""

import sys

from PySide6.QtWidgets import QApplication

from . import theme
from .main_window import MainWindow


def main(argv=None):
    app = QApplication(argv if argv is not None else sys.argv)
    app.setApplicationName("NVict Reader")
    app.setOrganizationName("NVict Service")

    theme.apply_theme(app, "Systeemstandaard")

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
