# -*- coding: utf-8 -*-
"""Persistente venstergeometrie via QSettings.

Eigen store ("NVict Service" / "NVict Reader Qt"), niet het bestaande
settings.json van de tkinter-app - voorkomt cross-contaminatie zolang beide
apps naast elkaar bestaan.
"""

from PySide6.QtCore import QSettings


def get_settings() -> QSettings:
    return QSettings("NVict Service", "NVict Reader Qt")


def save_window_state(window):
    settings = get_settings()
    settings.setValue("geometry", window.saveGeometry())
    settings.setValue("maximized", window.isMaximized())


def restore_window_state(window):
    settings = get_settings()
    geometry = settings.value("geometry")
    if geometry is not None:
        window.restoreGeometry(geometry)
    if settings.value("maximized", False, type=bool):
        window.showMaximized()
