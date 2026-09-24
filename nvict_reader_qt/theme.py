# -*- coding: utf-8 -*-
"""Thema-kleuren, overgenomen van NVict_Reader.py Theme-klasse (regel 487-505).

De hex-waarden zijn ongewijzigd voor een consistente merk-look; alleen de
toepassing verandert van losse tkinter-widgetkleuren naar QPalette + QSS.
"""

from PySide6.QtCore import QEvent, QObject, Qt
from PySide6.QtGui import QPalette, QColor
from PySide6.QtWidgets import QApplication, QMenu, QStyleFactory

from .icon_utils import get_scrollbar_arrow_path
from .resources import get_resource_path

LIGHT = {
    "BG_PRIMARY": "#f3f3f3", "BG_SECONDARY": "#ffffff",
    "TEXT_PRIMARY": "#1c1c1c", "TEXT_SECONDARY": "#737373",
    "ACCENT_COLOR": "#10a2dd", "SUCCESS_COLOR": "#28a745",
    "WARNING_COLOR": "#ff8c00", "ERROR_COLOR": "#d13438",
    "SELECTION_COLOR": "#FFD700", "HOVER_COLOR": "#f5f5f5",
}
DARK = {
    "BG_PRIMARY": "#1e1e1e", "BG_SECONDARY": "#2d2d2d",
    "TEXT_PRIMARY": "#f0f0f0", "TEXT_SECONDARY": "#a0a0a0",
    "ACCENT_COLOR": "#10a2dd", "SUCCESS_COLOR": "#28a745",
    "WARNING_COLOR": "#ff8c00", "ERROR_COLOR": "#d13438",
    "SELECTION_COLOR": "#FFD700", "HOVER_COLOR": "#292929",
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
    dark = colors is DARK
    # Zelfde chevron-icoontjes als Vorige/Volgende pagina (testfeedback:
    # de standaard Fusion-pijltjes op de scrollbar zagen er niet uit) - in
    # donker thema geïnverteerd via get_scrollbar_arrow_path.
    up_arrow = get_scrollbar_arrow_path(get_resource_path("icons/prev-page.png"), dark).replace("\\", "/")
    down_arrow = get_scrollbar_arrow_path(get_resource_path("icons/next-page.png"), dark).replace("\\", "/")
    return f"""
        QMainWindow, QWidget {{
            background-color: {colors["BG_PRIMARY"]};
            color: {colors["TEXT_PRIMARY"]};
            font-family: "{FONT_FAMILY}";
        }}
        QToolBar {{
            background-color: {colors["BG_SECONDARY"]};
            border: none;
            spacing: 0px;
            padding: 2px;
        }}
        QToolBar::separator {{
            background-color: {colors["TEXT_SECONDARY"]};
            width: 1px;
            margin: 6px 6px;
        }}
        QToolButton {{
            background-color: transparent;
            border: none;
            border-radius: 4px;
            padding: 4px 6px;
            margin: 1px;
        }}
        QToolButton:hover, QToolButton:pressed {{
            background-color: {colors["HOVER_COLOR"]};
        }}
        QToolButton:checked {{
            background-color: {colors["HOVER_COLOR"]};
            border: 1px solid {colors["ACCENT_COLOR"]};
        }}
        QToolButton:disabled {{
            color: {colors["TEXT_SECONDARY"]};
        }}
        QToolButton::menu-indicator {{
            image: none;
            width: 0px;
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
            border: 1px solid {colors["TEXT_SECONDARY"]};
            border-radius: 8px;
            padding: 6px;
        }}
        QMenu::item {{
            border-radius: 5px;
            padding: 6px 24px 6px 12px;
            margin: 1px 0px;
        }}
        QMenu::item:selected {{
            background-color: {colors["ACCENT_COLOR"]};
            color: #ffffff;
        }}
        QMenu::item:disabled {{
            color: {colors["TEXT_SECONDARY"]};
        }}
        QMenu::separator {{
            height: 1px;
            background: {colors["TEXT_SECONDARY"]};
            margin: 6px 10px;
        }}
        QMenu::icon {{
            padding-left: 6px;
        }}
        QGroupBox {{
            border: 1px solid {colors["TEXT_SECONDARY"]};
            border-radius: 4px;
            margin-top: 10px;
            padding-top: 10px;
            font-weight: bold;
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            left: 8px;
            padding: 0 4px;
        }}
        QRadioButton::indicator, QCheckBox::indicator {{
            width: 14px;
            height: 14px;
            border: 2px solid {colors["TEXT_SECONDARY"]};
            background-color: {colors["BG_SECONDARY"]};
        }}
        QRadioButton::indicator {{
            border-radius: 8px;
        }}
        QCheckBox::indicator {{
            border-radius: 3px;
        }}
        QRadioButton::indicator:checked, QCheckBox::indicator:checked {{
            border: 2px solid {colors["ACCENT_COLOR"]};
            background-color: {colors["ACCENT_COLOR"]};
        }}
        QRadioButton::indicator:hover, QCheckBox::indicator:hover {{
            border-color: {colors["ACCENT_COLOR"]};
        }}
        QLineEdit, QPlainTextEdit, QTextEdit, QSpinBox, QComboBox {{
            background-color: {colors["BG_SECONDARY"]};
            color: {colors["TEXT_PRIMARY"]};
            border: 1px solid {colors["TEXT_SECONDARY"]};
            border-radius: 3px;
            padding: 2px 4px;
        }}
        QLineEdit:focus, QPlainTextEdit:focus, QTextEdit:focus, QSpinBox:focus, QComboBox:focus {{
            border: 1px solid {colors["ACCENT_COLOR"]};
        }}
        QScrollBar:vertical {{
            background: {colors["BG_PRIMARY"]};
            width: 16px;
            margin: 0px;
        }}
        QScrollBar::handle:vertical {{
            background: {colors["TEXT_SECONDARY"]};
            border-radius: 6px;
            min-height: 24px;
            margin: 2px;
        }}
        QScrollBar::handle:vertical:hover {{
            background: {colors["ACCENT_COLOR"]};
        }}
        QScrollBar::sub-line:vertical {{
            height: 16px;
            subcontrol-position: top;
            subcontrol-origin: margin;
            background: {colors["BG_SECONDARY"]};
        }}
        QScrollBar::add-line:vertical {{
            height: 16px;
            subcontrol-position: bottom;
            subcontrol-origin: margin;
            background: {colors["BG_SECONDARY"]};
        }}
        QScrollBar::sub-line:vertical:hover, QScrollBar::add-line:vertical:hover {{
            background: {colors["HOVER_COLOR"]};
        }}
        QScrollBar::up-arrow:vertical {{
            image: url({up_arrow});
            width: 10px;
            height: 10px;
        }}
        QScrollBar::down-arrow:vertical {{
            image: url({down_arrow});
            width: 10px;
            height: 10px;
        }}
        QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
            background: none;
        }}
        QScrollBar:horizontal {{
            background: {colors["BG_PRIMARY"]};
            height: 16px;
            margin: 0px;
        }}
        QScrollBar::handle:horizontal {{
            background: {colors["TEXT_SECONDARY"]};
            border-radius: 6px;
            min-width: 24px;
            margin: 2px;
        }}
        QScrollBar::handle:horizontal:hover {{
            background: {colors["ACCENT_COLOR"]};
        }}
        QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
            background: none;
        }}
    """


class _RoundedMenuFilter(QObject):
    """Maakt het venster achter élk QMenu doorzichtig en randloos, zodat de
    afgeronde hoeken uit de QSS (border-radius) echt rond zijn.

    Alleen WA_TranslucentBackground is op Windows niet genoeg: het popup-
    venster houdt dan zijn eigen rechthoekige rand/schaduw, die als zwart
    vlakje buiten de afgeronde hoek zichtbaar blijft. Via een app-breed
    event-filter (op Polish, vóór het eerste tonen) geldt dit ook voor
    menu's die Qt zelf maakt, zoals het rechtermuisklikmenu van tekstvelden.
    """

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.Polish and isinstance(obj, QMenu) and not obj.property("_nv_rounded"):
            obj.setProperty("_nv_rounded", True)
            obj.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
            obj.setWindowFlag(Qt.WindowType.FramelessWindowHint, True)
            obj.setWindowFlag(Qt.WindowType.NoDropShadowWindowHint, True)
        return False


_menu_filter = None


def _install_menu_filter(app: QApplication):
    global _menu_filter
    if _menu_filter is None:
        _menu_filter = _RoundedMenuFilter(app)
        app.installEventFilter(_menu_filter)


def apply_theme(app: QApplication, mode: str = "Systeemstandaard") -> dict:
    """Pas het thema toe op de hele applicatie en geef de gebruikte kleuren terug.

    Forceert de Fusion-stijl: Windows' eigen stijl ("windowsvista"/
    "windows11") negeert een deel van onze QSS (bv. de scrollbar-pijltjes
    en de vlakke QToolButton-hover-kleur) omdat die stijl scrollbars/
    knoppen native tekent - Fusion respecteert QSS/QPalette volledig en
    geeft zo een consistent uiterlijk in zowel licht als donker thema.
    """
    app.setStyle(QStyleFactory.create("Fusion"))
    _install_menu_filter(app)
    colors = resolve_theme(mode)
    app.setPalette(build_palette(colors))
    app.setStyleSheet(build_stylesheet(colors))
    return colors
