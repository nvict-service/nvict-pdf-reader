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


def get_theme_mode() -> str:
    return get_settings().value("theme", "Systeemstandaard", type=str)


def save_theme_mode(mode: str):
    get_settings().setValue("theme", mode)


MAX_RECENT_FILES = 10


def get_recent_files() -> list:
    """Poort van add_to_recent_files/recent_files uit NVict_Reader.py (regel
    1680-1692) - meest recent eerst, max MAX_RECENT_FILES bestanden."""
    recent = get_settings().value("recent_files", [], type=list)
    return [path for path in recent if path]


def add_recent_file(file_path: str):
    recent = get_recent_files()
    if file_path in recent:
        recent.remove(file_path)
    recent.insert(0, file_path)
    get_settings().setValue("recent_files", recent[:MAX_RECENT_FILES])


def get_show_thumbnails_default() -> bool:
    """Of het paginaminiaturen-paneel standaard getoond wordt bij het
    openen van een document - door de gebruiker instelbaar (Instellingen)."""
    return get_settings().value("show_thumbnails_default", True, type=bool)


def save_show_thumbnails_default(value: bool):
    get_settings().setValue("show_thumbnails_default", value)
