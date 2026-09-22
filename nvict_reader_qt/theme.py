# -*- coding: utf-8 -*-
"""Thema-kleuren, overgenomen van NVict_Reader.py Theme-klasse (regel 487-505).

De hex-waarden zijn ongewijzigd voor een consistente merk-look; alleen de
toepassing verandert van losse tkinter-widgetkleuren naar QPalette + QSS.
"""

from PySide6.QtGui import QPalette, QColor
from PySide6.QtWidgets import QApplication

LIGHT = {
    "BG_PRIMARY": "#f3f3f3", "BG_SECONDARY": "#ffffff",
    "TEXT_PRIMARY": "#1c1c1c", "TEXT_SECONDARY": "#737373",
    "ACCENT_COLOR": "#10a2dd", "SUCCESS_COLOR": "#28a745",
    "WARNING_COLOR": "#ff8c00", "ERROR_COLOR": "#d13438",
    "SELECTION_COLOR": "#FFD700",
}
DARK = {
    "BG_PRIMARY": "#1e1e1e", "BG_SECONDARY": "#2d2d2d",
    "TEXT_PRIMARY": "#f0f0f0", "TEXT_SECONDARY": "#a0a0a0",
    "ACCENT_COLOR": "#10a2dd", "SUCCESS_COLOR": "#28a745",
    "WARNING_COLOR": "#ff8c00", "ERROR_COLOR": "#d13438",
    "SELECTION_COLOR": "#FFD700",
}

FONT_FAMILY = "Segoe UI Variable"


def _system_is_dark():
    """Bepaal of het Windows-systeemthema donker is."""
    try:
        style_hints = QApplication.instance().styleHints()
        scheme = style_hints.colorScheme()
        # Qt.ColorScheme.Dark == 2, Light == 1, Unknown == 0 (Qt 6.5+)
        return scheme.name == "Dark"
    except Exception:
        return False


def resolve_theme(mode: str) -> dict:
    """Zet een thema-modus ("Licht"/"Donker"/"Systeemstandaard") om naar een kleurendict."""
    if mode == "Licht":
        return LIGHT
    if mode == "Donker":
        return DARK
    return DARK if _system_is_dark() else LIGHT


def build_palette(colors: dict) -> QPalette:
    """Bouw een QPalette met de basiskleuren die Qt-widgets zelf gebruiken."""
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(colors["BG_PRIMARY"]))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(colors["TEXT_PRIMARY"]))
    palette.setColor(QPalette.ColorRole.Base, QColor(colors["BG_SECONDARY"]))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(colors["BG_PRIMARY"]))
    palette.setColor(QPalette.ColorRole.Text, QColor(colors["TEXT_PRIMARY"]))
    palette.setColor(QPalette.ColorRole.Button, QColor(colors["BG_SECONDARY"]))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(colors["TEXT_PRIMARY"]))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(colors["ACCENT_COLOR"]))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(colors["BG_SECONDARY"]))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor(colors["TEXT_PRIMARY"]))
    return palette


def build_stylesheet(colors: dict) -> str:
    """Bouw een QSS-stylesheet voor widgets waar QPalette te grof is."""
    return f"""
        QMainWindow, QWidget {{
            background-color: {colors["BG_PRIMARY"]};
            color: {colors["TEXT_PRIMARY"]};
            font-family: "{FONT_FAMILY}";
        }}
        QToolBar {{
            background-color: {colors["BG_SECONDARY"]};
            border: none;
            spacing: 4px;
        }}
        QTabWidget::pane {{
            border: none;
        }}
        QTabBar::tab {{
            background-color: {colors["BG_PRIMARY"]};
            color: {colors["TEXT_SECONDARY"]};
            padding: 6px 12px;
        }}
        QTabBar::tab:selected {{
            background-color: {colors["BG_SECONDARY"]};
            color: {colors["TEXT_PRIMARY"]};
        }}
        QGraphicsView {{
            background-color: {colors["BG_PRIMARY"]};
            border: none;
        }}
        QMenuBar {{
            background-color: {colors["BG_SECONDARY"]};
            color: {colors["TEXT_PRIMARY"]};
        }}
        QMenuBar::item:selected {{
            background-color: {colors["ACCENT_COLOR"]};
            color: #ffffff;
        }}
        QMenu {{
            background-color: {colors["BG_SECONDARY"]};
            color: {colors["TEXT_PRIMARY"]};
        }}
        QMenu::item:selected {{
            background-color: {colors["ACCENT_COLOR"]};
            color: #ffffff;
        }}
    """


def apply_theme(app: QApplication, mode: str = "Systeemstandaard") -> dict:
    """Pas het thema toe op de hele applicatie en geef de gebruikte kleuren terug."""
    colors = resolve_theme(mode)
    app.setPalette(build_palette(colors))
    app.setStyleSheet(build_stylesheet(colors))
    return colors
