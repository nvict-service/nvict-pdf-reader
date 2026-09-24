# -*- coding: utf-8 -*-
"""Hoofdvenster: menu, toolbar en tabbladen voor meerdere open PDF's.

Bevat zelf geen documentstate - delegeert naar de actieve tab
(PdfTabWidget), analoog aan hoe NVictReader in de tkinter-versie delegeert
naar self.get_active_tab().
"""

import os
import webbrowser
from datetime import datetime

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QIcon, QIntValidator, QKeySequence, QPageLayout, QPainter, QPalette, QPixmap
from PySide6.QtPrintSupport import QPrinter
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressDialog,
    QStackedWidget,
    QTabWidget,
    QToolButton,
)

from . import document_tools, email_sender, i18n, print_backend, save_pdf, settings, theme, update_checker
from .about_dialog import AboutDialog
from .document_tools_dialogs import ExportPagesDialog, MergePdfsDialog, RotatePagesDialog
from .fullscreen_view import FullscreenWindow
from .i18n import tr
from .icon_utils import invert_icon_colors
from .pdf_tab import PdfTabWidget
from .print_dialog import PrintDialog
from .resources import get_resource_path
from .search_dialog import SearchDialog
from .settings_dialog import SettingsDialog
from .welcome_widget import WelcomeWidget

SOFTWARE_URL = "https://nvict.nl/software-download"


def _is_dark_theme() -> bool:
    app = QApplication.instance()
    if app is None:
        return False
    return app.palette().color(QPalette.ColorRole.Window).lightness() < 128


def _faded_pixmap(pixmap: QPixmap, opacity: float = 0.35) -> QPixmap:
    """Maak een duidelijk uitgegrijsde variant voor QIcon.Mode.Disabled.

    Qt's automatisch gegenereerde disabled-icoon is op een donkere
    werkbalk te subtiel om te zien (testfeedback: onduidelijk welke
    knoppen momenteel niets doen) - deze variant is altijd zichtbaar
    minder opvallend, in elk thema.
    """
    faded = QPixmap(pixmap.size())
    faded.setDevicePixelRatio(pixmap.devicePixelRatio())
    faded.fill(Qt.GlobalColor.transparent)
    painter = QPainter(faded)
    painter.setOpacity(opacity)
    painter.drawPixmap(0, 0, pixmap)
    painter.end()
    return faded


def _icon(name: str) -> QIcon:
    path = get_resource_path(os.path.join("icons", name))
    if not os.path.exists(path):
        return QIcon()
    pixmap = QPixmap(path)
    if _is_dark_theme():
        pixmap = invert_icon_colors(pixmap)
    icon = QIcon()
    icon.addPixmap(pixmap, QIcon.Mode.Normal)
    icon.addPixmap(_faded_pixmap(pixmap), QIcon.Mode.Disabled)
    return icon


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("NVict Reader")
        self.resize(1366, 768)
        self.setMinimumSize(800, 600)
        favicon_path = get_resource_path("favicon.ico")
        if os.path.exists(favicon_path):
            self.setWindowIcon(QIcon(favicon_path))

        self._thumbnails_visible = settings.get_show_thumbnails_default()
        self._fullscreen_window = None

        self.tabs = QTabWidget(self)
        self.tabs.setTabsClosable(True)
        self.tabs.setMovable(True)
        self.tabs.tabCloseRequested.connect(self._close_tab)
        self.tabs.currentChanged.connect(self._on_tab_changed)

        self.welcome_widget = WelcomeWidget(self)
        self.welcome_widget.file_chosen.connect(self.open_file)

        self.central_stack = QStackedWidget(self)
        self.central_stack.addWidget(self.welcome_widget)
        self.central_stack.addWidget(self.tabs)
        self.setCentralWidget(self.central_stack)

        self._build_actions()
        self._build_menu()
        self._build_toolbar()
        self._build_status_bar()
        self._update_actions_enabled()
        self._update_central_widget()

        settings.restore_window_state(self)

    # ── Actions ──────────────────────────────────────────────────────

    def _build_actions(self):
        self.action_open = QAction(_icon("open.png"), tr("&Openen..."), self)
        self.action_open.setShortcut(QKeySequence.StandardKey.Open)
        self.action_open.triggered.connect(self.open_file_dialog)

        self.action_close_tab = QAction(tr("Tab &sluiten"), self)
        self.action_close_tab.setShortcut("Ctrl+W")
        self.action_close_tab.triggered.connect(lambda: self._close_tab(self.tabs.currentIndex()))

        self.action_quit = QAction(tr("&Afsluiten"), self)
        self.action_quit.setShortcut(QKeySequence.StandardKey.Quit)
        self.action_quit.triggered.connect(self.close)

        self.action_zoom_in = QAction(_icon("zoom-in.png"), tr("Zoom &in"), self)
        self.action_zoom_in.setShortcut("Ctrl+=")
        self.action_zoom_in.triggered.connect(lambda: self._on_active(lambda v: v.zoom_in()))

        self.action_zoom_out = QAction(_icon("zoom-out.png"), tr("Zoom &uit"), self)
        self.action_zoom_out.setShortcut("Ctrl+-")
        self.action_zoom_out.triggered.connect(lambda: self._on_active(lambda v: v.zoom_out()))

        self.action_fit_width = QAction(_icon("fit-width.png"), tr("&Pasbreedte"), self)
        self.action_fit_width.setShortcut("Ctrl+0")
        self.action_fit_width.triggered.connect(lambda: self._on_active(lambda v: v.set_zoom_mode_fit_width()))

        self.action_first_page = QAction(_icon("first-page.png"), tr("&Eerste pagina"), self)
        self.action_first_page.setShortcut("Ctrl+Home")
        self.action_first_page.triggered.connect(lambda: self._on_active(lambda v: v.first_page()))

        self.action_prev_page = QAction(_icon("prev-page.png"), tr("&Vorige pagina"), self)
        self.action_prev_page.setShortcut(QKeySequence.StandardKey.MoveToPreviousPage)
        self.action_prev_page.triggered.connect(lambda: self._on_active(lambda v: v.prev_page()))

        self.action_next_page = QAction(_icon("next-page.png"), tr("&Volgende pagina"), self)
        self.action_next_page.setShortcut(QKeySequence.StandardKey.MoveToNextPage)
        self.action_next_page.triggered.connect(lambda: self._on_active(lambda v: v.next_page()))

        self.action_last_page = QAction(_icon("last-page.png"), tr("&Laatste pagina"), self)
        self.action_last_page.setShortcut("Ctrl+End")
        self.action_last_page.triggered.connect(lambda: self._on_active(lambda v: v.last_page()))

        self.action_print = QAction(_icon("print.png"), tr("&Afdrukken..."), self)
        self.action_print.setShortcut(QKeySequence.StandardKey.Print)
        self.action_print.triggered.connect(self._print_current)

        self.action_save_as = QAction(_icon("save.png"), tr("&Opslaan als..."), self)
        self.action_save_as.setShortcut(QKeySequence.StandardKey.Save)
        self.action_save_as.triggered.connect(self._save_as_current)

        # "Geen hulpmiddel": expliciete, altijd-zichtbare manier om een
        # actief hulpmiddel weer uit te zetten - zonder dit was de enige
        # manier om te stoppen het (niet voor de hand liggende) opnieuw
        # aanklikken van hetzelfde, allang-aangevinkte menu-item.
        self.action_select_tool = QAction(_icon("close.png"), tr("&Geen hulpmiddel (Esc)"), self)
        self.action_select_tool.setCheckable(True)
        self.action_select_tool.setChecked(True)
        self.action_select_tool.toggled.connect(self._on_select_tool_toggled)

        # Onzichtbare actie enkel voor de Escape-sneltoets - los van het
        # zichtbare menu-item, zodat hem indrukken terwijl al geen
        # hulpmiddel actief is geen ongewenste toggle veroorzaakt.
        self.action_escape_tool = QAction(self)
        self.action_escape_tool.setShortcut(QKeySequence(Qt.Key.Key_Escape))
        self.action_escape_tool.triggered.connect(self._deactivate_tools)
        self.addAction(self.action_escape_tool)

        self.action_text_annotate = QAction(_icon("type-text.png"), tr("&Tekst toevoegen"), self)
        self.action_text_annotate.setCheckable(True)
        self.action_text_annotate.toggled.connect(self._on_text_annotate_toggled)

        self.action_highlight = QAction(_icon("marker.png"), tr("&Markeren"), self)
        self.action_highlight.setCheckable(True)
        self.action_highlight.toggled.connect(self._on_highlight_toggled)

        self.action_form_mode = QAction(_icon("form.png"), tr("&Formulier invullen"), self)
        self.action_form_mode.setCheckable(True)
        self.action_form_mode.toggled.connect(self._on_form_mode_toggled)

        self.action_signature = QAction(_icon("check.png"), tr("&Handtekening plaatsen"), self)
        self.action_signature.setCheckable(True)
        self.action_signature.toggled.connect(self._on_signature_toggled)

        self.action_export_pages = QAction(_icon("pdf.png"), tr("Pagina's &exporteren..."), self)
        self.action_export_pages.triggered.connect(self._export_pages_current)

        self.action_merge_pdfs = QAction(_icon("copy.png"), tr("PDF's &samenvoegen..."), self)
        self.action_merge_pdfs.triggered.connect(self._merge_pdfs_current)

        self.action_rotate_pages = QAction(_icon("reset.png"), tr("Pagina &roteren..."), self)
        self.action_rotate_pages.triggered.connect(self._rotate_pages_current)

        self.action_copy_text = QAction(_icon("copy.png"), tr("&Kopiëren"), self)
        self.action_copy_text.setShortcut(QKeySequence.StandardKey.Copy)
        self.action_copy_text.triggered.connect(lambda: self._on_active(lambda v: v.copy_selected_text()))

        self.action_settings = QAction(tr("&Instellingen..."), self)
        self.action_settings.triggered.connect(self._open_settings)

        self.action_check_updates = QAction(tr("Controleren op &updates..."), self)
        self.action_check_updates.triggered.connect(lambda: update_checker.check_for_updates(self, silent=False))

        self.action_pdf_info = QAction(tr("&PDF-informatie..."), self)
        self.action_pdf_info.triggered.connect(self._show_pdf_info)

        self.action_about = QAction(tr("&Over NVict Reader..."), self)
        self.action_about.triggered.connect(self._show_about)

        # "book.png" (paginaminiaturen -> pages.png) komt hierdoor vrij voor
        # de nieuwe boek-modus-knop hieronder.
        self.action_toggle_thumbnails = QAction(_icon("pages.png"), tr("&Pagina's"), self)
        self.action_toggle_thumbnails.setCheckable(True)
        self.action_toggle_thumbnails.setChecked(self._thumbnails_visible)
        self.action_toggle_thumbnails.toggled.connect(self._on_toggle_thumbnails)

        self.action_send = QAction(_icon("send.png"), tr("&Verzenden..."), self)
        self.action_send.triggered.connect(self._send_current)

        # Hergebruikt reset.png (cirkelpijl) - past semantisch net zo goed
        # bij "ongedaan maken" als bij "roteren", zelfde icoon-hergebruik-
        # patroon als copy.png elders in deze dict.
        self.action_undo = QAction(_icon("reset.png"), tr("&Ongedaan maken"), self)
        self.action_undo.setShortcut("Ctrl+Z")
        self.action_undo.triggered.connect(lambda: self._on_active(lambda v: v.undo()))

        self.action_search = QAction(_icon("search.png"), tr("&Zoeken..."), self)
        self.action_search.setShortcut("Ctrl+F")
        self.action_search.triggered.connect(self._open_search)

        self.action_book_mode = QAction(_icon("book.png"), tr("&Boekweergave"), self)
        self.action_book_mode.setCheckable(True)
        self.action_book_mode.toggled.connect(lambda checked: self._on_active(lambda v: v.toggle_book_mode()))

        self.action_fullscreen = QAction(_icon("full-screen.png"), tr("&Volledig scherm"), self)
        self.action_fullscreen.setShortcut("F11")
        self.action_fullscreen.triggered.connect(self._enter_fullscreen)

        # Onthouden welk icoonbestand bij welke actie hoort, zodat na een
        # thema-wissel alle iconen opnieuw geladen (en zo nodig opnieuw
        # geïnverteerd) kunnen worden - zie _refresh_icons.
        self._icon_actions = {
            self.action_open: "open.png", self.action_zoom_in: "zoom-in.png",
            self.action_zoom_out: "zoom-out.png", self.action_fit_width: "fit-width.png",
            self.action_first_page: "first-page.png", self.action_prev_page: "prev-page.png",
            self.action_next_page: "next-page.png", self.action_last_page: "last-page.png",
            self.action_print: "print.png", self.action_save_as: "save.png",
            self.action_text_annotate: "type-text.png", self.action_highlight: "marker.png",
            self.action_form_mode: "form.png", self.action_signature: "check.png",
            self.action_select_tool: "close.png",
            self.action_merge_pdfs: "copy.png",
            self.action_rotate_pages: "reset.png", self.action_copy_text: "copy.png",
            self.action_toggle_thumbnails: "pages.png", self.action_send: "send.png",
            self.action_search: "search.png", self.action_book_mode: "book.png",
            self.action_fullscreen: "full-screen.png", self.action_undo: "reset.png",
            self.action_export_pages: "pdf.png",
        }

    def _refresh_icons(self):
        """Herlaad alle actie-iconen - nodig na een thema-wissel zodat de
        donker-thema-inversie (zie icon_utils.py) opnieuw wordt toegepast."""
        for action, name in self._icon_actions.items():
            action.setIcon(_icon(name))
        self.edit_menu_button.setIcon(_icon("toolbox.png"))
        self.tools_menu_button.setIcon(_icon("toolbox.png"))
        self._style_toolbar_labels()

    def _style_toolbar_labels(self):
        """Forceer de tekstkleur van kale QLabels op de werkbalk.

        Fusion geeft een los met `addWidget` toegevoegde QLabel een eigen
        "WindowText"-kleur die niet meeloopt met de rest van het lint
        (bleef wit/grijs, los van het thema) - expliciet zetten lost dat
        op, in plaats van te vertrouwen op QSS-overerving die hier niet
        werkt zoals bij QToolButton/QAction.
        """
        colors = theme.resolve_theme(settings.get_theme_mode())
        self.page_count_label.setStyleSheet(f"color: {colors['TEXT_PRIMARY']}; background: transparent;")

    def _build_menu(self):
        file_menu = self.menuBar().addMenu(tr("&Bestand"))
        file_menu.addAction(self.action_open)
        file_menu.addAction(self.action_close_tab)
        file_menu.addSeparator()
        file_menu.addAction(self.action_print)
        file_menu.addAction(self.action_save_as)
        file_menu.addAction(self.action_send)
        file_menu.addSeparator()
        file_menu.addAction(self.action_undo)
        file_menu.addAction(self.action_copy_text)
        file_menu.addSeparator()
        file_menu.addAction(self.action_quit)

        view_menu = self.menuBar().addMenu(tr("&Beeld"))
        view_menu.addAction(self.action_toggle_thumbnails)
        view_menu.addAction(self.action_book_mode)
        view_menu.addAction(self.action_fullscreen)
        view_menu.addSeparator()
        view_menu.addAction(self.action_search)
        view_menu.addSeparator()
        view_menu.addAction(self.action_zoom_in)
        view_menu.addAction(self.action_zoom_out)
        view_menu.addAction(self.action_fit_width)
        view_menu.addSeparator()
        view_menu.addAction(self.action_first_page)
        view_menu.addAction(self.action_prev_page)
        view_menu.addAction(self.action_next_page)
        view_menu.addAction(self.action_last_page)

        tools_menu = self.menuBar().addMenu(tr("&Hulpmiddelen"))
        tools_menu.addAction(self.action_select_tool)
        tools_menu.addSeparator()
        tools_menu.addAction(self.action_text_annotate)
        tools_menu.addAction(self.action_highlight)
        tools_menu.addAction(self.action_form_mode)
        tools_menu.addAction(self.action_signature)
        self.tools_menu = tools_menu

        edit_menu = self.menuBar().addMenu(tr("&Bewerken"))
        edit_menu.addAction(self.action_export_pages)
        edit_menu.addAction(self.action_merge_pdfs)
        edit_menu.addAction(self.action_rotate_pages)
        self.edit_menu = edit_menu

        settings_menu = self.menuBar().addMenu(tr("&Instellingen"))
        settings_menu.addAction(self.action_settings)

        help_menu = self.menuBar().addMenu(tr("&Help"))
        help_menu.addAction(self.action_pdf_info)
        help_menu.addSeparator()
        help_menu.addAction(self.action_check_updates)
        help_menu.addSeparator()
        help_menu.addAction(self.action_about)
        # Afgeronde hoeken van alle menu's: zie theme._RoundedMenuFilter.

    # Onder deze breedte past het volledige lint met tekst-onder-icoon niet
    # meer (natuurlijke breedte ligt rond de 3200px) - dan schakelen we over
    # op alleen-icoon-knoppen (~1050px), zodat alle knoppen zichtbaar
    # blijven zonder dat je de Qt-eigen overloop-pijl (">>") moet ontdekken.
    _TOOLBAR_COMPACT_WIDTH = 1300

    def _build_toolbar(self):
        toolbar = self.addToolBar(tr("Hoofdwerkbalk"))
        self._toolbar = toolbar
        self._toolbar_compact = False
        toolbar.setMovable(False)
        toolbar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        toolbar.addAction(self.action_open)
        toolbar.addAction(self.action_print)
        toolbar.addAction(self.action_save_as)
        toolbar.addAction(self.action_send)
        toolbar.addAction(self.action_search)
        toolbar.addSeparator()
        toolbar.addAction(self.action_first_page)
        toolbar.addAction(self.action_prev_page)
        toolbar.addAction(self.action_next_page)
        toolbar.addAction(self.action_last_page)

        # Platte tekstinvoer + Enter, net als de oude tkinter-versie (géén
        # spinner-knoppen: die navigeerden pas na focusverlies, niet direct
        # bij een klik, wat verwarrend aanvoelde - zie testfeedback).
        self.page_goto_edit = QLineEdit(toolbar)
        self.page_goto_edit.setFixedWidth(40)
        self.page_goto_edit.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.page_goto_edit.setValidator(QIntValidator(1, 999999, self))
        self.page_goto_edit.setToolTip(tr("Ga naar pagina"))
        self.page_goto_edit.returnPressed.connect(self._on_page_goto_changed)
        toolbar.addWidget(self.page_goto_edit)
        self.page_count_label = QLabel("/ 0", toolbar)
        self.page_count_label.setContentsMargins(4, 0, 8, 0)
        toolbar.addWidget(self.page_count_label)

        toolbar.addSeparator()
        toolbar.addAction(self.action_toggle_thumbnails)
        toolbar.addAction(self.action_book_mode)
        toolbar.addSeparator()
        toolbar.addAction(self.action_zoom_out)
        toolbar.addAction(self.action_zoom_in)
        toolbar.addAction(self.action_fit_width)
        toolbar.addAction(self.action_fullscreen)
        toolbar.addSeparator()
        toolbar.addAction(self.action_undo)
        toolbar.addAction(self.action_copy_text)
        toolbar.addSeparator()

        self._style_toolbar_labels()

        self.tools_menu_button = QToolButton(toolbar)
        self.tools_menu_button.setIcon(_icon("toolbox.png"))
        self.tools_menu_button.setText(tr("Hulpmiddelen") + " ▼")
        self.tools_menu_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        self.tools_menu_button.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.tools_menu_button.setMenu(self.tools_menu)
        toolbar.addWidget(self.tools_menu_button)

        self.edit_menu_button = QToolButton(toolbar)
        self.edit_menu_button.setIcon(_icon("toolbox.png"))
        self.edit_menu_button.setText(tr("Bewerken") + " ▼")
        self.edit_menu_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        self.edit_menu_button.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.edit_menu_button.setMenu(self.edit_menu)
        toolbar.addWidget(self.edit_menu_button)

    def _update_toolbar_style(self):
        """Schakelt tussen tekst-onder-icoon (breed venster) en alleen-icoon
        (smal venster) - de knoptekst kost het meeste van de ~3200px die het
        volledige lint nodig heeft; zonder tekst is dat maar ~1050px, wat op
        vrijwel elk scherm past. De tooltips blijven de knoptekst tonen."""
        compact = self.width() < self._TOOLBAR_COMPACT_WIDTH
        if compact == self._toolbar_compact:
            return
        self._toolbar_compact = compact
        style = (
            Qt.ToolButtonStyle.ToolButtonIconOnly
            if compact
            else Qt.ToolButtonStyle.ToolButtonTextUnderIcon
        )
        self._toolbar.setToolButtonStyle(style)
        self.tools_menu_button.setToolButtonStyle(style)
        self.edit_menu_button.setToolButtonStyle(style)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_toolbar_style()

    def _build_status_bar(self):
        status_bar = self.statusBar()

        self.doc_info_label = QLabel(tr("Geen document geopend"), status_bar)
        status_bar.addWidget(self.doc_info_label)

        copyright_label = QLabel(f"© {datetime.now().year} NVict Service - www.nvict.nl", status_bar)
        copyright_label.setCursor(Qt.CursorShape.PointingHandCursor)
        copyright_label.mousePressEvent = lambda _event: webbrowser.open(SOFTWARE_URL)
        status_bar.addPermanentWidget(copyright_label)

    def _update_doc_info_label(self):
        tab = self.tabs.currentWidget()
        if tab is None or not tab.view.pdf_document:
            self.doc_info_label.setText(tr("Geen document geopend"))
            return
        text = tr("{title}  —  pagina {page} / {count}", title=tab.title, page=tab.view.current_page + 1, count=tab.page_count)
        if tab.view.security_info:
            text += f"   ·   {tab.view.security_info}"
        self.doc_info_label.setText(text)

    def _update_central_widget(self):
        if self.tabs.count() == 0:
            self.welcome_widget.refresh()
            self.central_stack.setCurrentWidget(self.welcome_widget)
        else:
            self.central_stack.setCurrentWidget(self.tabs)

    def _update_actions_enabled(self):
        has_tab = self.tabs.count() > 0
        tab = self.tabs.currentWidget()
        view = tab.view if tab is not None else None
        for action in (
            self.action_close_tab, self.action_zoom_in, self.action_zoom_out,
            self.action_fit_width, self.action_first_page, self.action_prev_page,
            self.action_next_page, self.action_last_page, self.action_print,
            self.action_text_annotate, self.action_highlight,
            self.action_signature, self.action_export_pages,
            self.action_merge_pdfs, self.action_rotate_pages, self.action_copy_text,
            self.action_toggle_thumbnails, self.action_search, self.action_book_mode,
            self.action_fullscreen, self.page_goto_edit, self.action_send,
            self.action_pdf_info,
        ):
            action.setEnabled(has_tab)
        self.edit_menu_button.setEnabled(has_tab)
        self.tools_menu_button.setEnabled(has_tab)
        self.action_save_as.setEnabled(has_tab and view is not None and view.has_unsaved_changes())
        self.action_form_mode.setEnabled(has_tab and view is not None and view._document_has_widgets())
        self.action_undo.setEnabled(has_tab and view is not None and view.can_undo())

    # ── Tab helpers ──────────────────────────────────────────────────

    def _on_active(self, func):
        tab = self.tabs.currentWidget()
        if tab is not None:
            func(tab.view)

    def _on_tab_changed(self, _index):
        self._update_actions_enabled()
        tab = self.tabs.currentWidget()
        mode = tab.view.tool_mode if tab is not None else None
        form_active = tab.view.form_mode if tab is not None else False
        actions = (self.action_text_annotate, self.action_highlight, self.action_form_mode, self.action_signature)
        for action in actions:
            action.blockSignals(True)
        self.action_text_annotate.setChecked(mode == "text_annotate")
        self.action_highlight.setChecked(mode == "highlight")
        self.action_signature.setChecked(mode == "signature")
        self.action_form_mode.setChecked(form_active)
        for action in actions:
            action.blockSignals(False)
        self._sync_select_tool_checked()

        self.action_book_mode.blockSignals(True)
        self.action_book_mode.setChecked(tab.view.book_mode if tab is not None else False)
        self.action_book_mode.blockSignals(False)

        self._sync_page_goto(tab)
        self._update_doc_info_label()

    def _uncheck_other_tools(self, keep):
        for action in (
            self.action_select_tool, self.action_text_annotate,
            self.action_highlight, self.action_form_mode, self.action_signature,
        ):
            if action is not keep:
                action.setChecked(False)

    def _on_select_tool_toggled(self, checked):
        if checked:
            self._uncheck_other_tools(self.action_select_tool)
        self._apply_tool_mode()

    def _deactivate_tools(self):
        """Escape-sneltoets: zet alle hulpmiddelen (inclusief formulier-
        invulmodus) direct uit, ongeacht wat er nu actief is."""
        self.action_select_tool.setChecked(True)

    def _on_text_annotate_toggled(self, checked):
        if checked:
            self._uncheck_other_tools(self.action_text_annotate)
        self._apply_tool_mode()

    def _on_highlight_toggled(self, checked):
        if checked:
            self._uncheck_other_tools(self.action_highlight)
        self._apply_tool_mode()

    def _on_signature_toggled(self, checked):
        if checked:
            self._uncheck_other_tools(self.action_signature)
        self._apply_tool_mode()

    def _apply_tool_mode(self):
        mode = None
        if self.action_text_annotate.isChecked():
            mode = "text_annotate"
        elif self.action_highlight.isChecked():
            mode = "highlight"
        elif self.action_signature.isChecked():
            mode = "signature"
        self._on_active(lambda v: v.set_tool_mode(mode))
        self._sync_select_tool_checked()

    def _sync_select_tool_checked(self):
        """Houd "Geen hulpmiddel" gelijk met de werkelijke staat - ook
        wanneer een tool-actie zelf (niet via _deactivate_tools) wordt
        uitgevinkt, bv. door er opnieuw op te klikken in het menu."""
        none_active = not (
            self.action_text_annotate.isChecked()
            or self.action_highlight.isChecked()
            or self.action_signature.isChecked()
            or self.action_form_mode.isChecked()
        )
        self.action_select_tool.blockSignals(True)
        self.action_select_tool.setChecked(none_active)
        self.action_select_tool.blockSignals(False)

    def _on_toggle_thumbnails(self, checked):
        self._thumbnails_visible = checked
        for index in range(self.tabs.count()):
            widget = self.tabs.widget(index)
            if widget is not None:
                widget.set_thumbnails_visible(checked)

    def _on_form_mode_toggled(self, checked):
        if checked:
            self.action_text_annotate.setChecked(False)
            self.action_highlight.setChecked(False)
            self.action_signature.setChecked(False)
            self.action_select_tool.setChecked(False)

        tab = self.tabs.currentWidget()
        if tab is None:
            self._sync_select_tool_checked()
            return

        if checked:
            if not tab.view.set_form_mode(True):
                self.action_form_mode.blockSignals(True)
                self.action_form_mode.setChecked(False)
                self.action_form_mode.blockSignals(False)
                QMessageBox.information(
                    self, tr("Geen formuliervelden"), tr("Dit document heeft geen invulbare formuliervelden.")
                )
        else:
            tab.view.set_form_mode(False)
        self._sync_select_tool_checked()

    def _close_tab(self, index):
        if index < 0:
            return
        tab = self.tabs.widget(index)
        if tab is not None and not save_pdf.confirm_discard_unsaved(self, tab.view):
            return
        self.tabs.removeTab(index)
        if tab is not None:
            tab.close_document()
            tab.deleteLater()
        self._update_actions_enabled()
        self._update_central_widget()

    # ── Ga naar pagina (toolbar) ─────────────────────────────────────

    def _sync_page_goto(self, tab):
        count = tab.page_count if tab is not None else 0
        self.page_goto_edit.setText(str((tab.view.current_page + 1) if tab is not None else 1))
        self.page_count_label.setText(f"/ {count}")

    def _on_page_goto_changed(self):
        """Poort van NVict_Reader.py:3055-3065 (go_to_page): bij een ongeldig
        of buiten-bereik paginanummer wordt het veld gewoon stilzwijgend
        teruggezet naar de huidige pagina, zonder foutmelding."""
        tab = self.tabs.currentWidget()
        if tab is None:
            return
        try:
            page_num = int(self.page_goto_edit.text()) - 1
        except ValueError:
            self._sync_page_goto(tab)
            return
        if 0 <= page_num < tab.page_count:
            tab.view.go_to_page(page_num)
        else:
            self._sync_page_goto(tab)

    def _on_page_changed_for_tab(self, tab, _page_num):
        if self.tabs.currentWidget() is tab:
            self._sync_page_goto(tab)
            self._update_doc_info_label()

    # ── Zoeken ────────────────────────────────────────────────────────

    def _open_search(self):
        tab = self.tabs.currentWidget()
        if tab is None or not tab.view.pdf_document:
            return
        dialog = SearchDialog(self)
        if not dialog.exec():
            return
        term = dialog.get_search_term()
        if not term:
            return
        if not tab.view.search(term):
            QMessageBox.information(self, tr("Zoeken"), tr("'{term}' niet gevonden in document", term=term))

    # ── Volledig scherm ──────────────────────────────────────────────

    def _enter_fullscreen(self):
        tab = self.tabs.currentWidget()
        if tab is None or not tab.view.pdf_document:
            QMessageBox.information(self, tr("Geen PDF"), tr("Open eerst een PDF om de presentatiemodus te gebruiken."))
            return

        def on_close(last_page, source_tab=tab):
            self._fullscreen_window = None
            source_tab.view.go_to_page(last_page)

        self._fullscreen_window = FullscreenWindow(tab.view.pdf_document, tab.view.current_page, on_close)

    # ── File handling ────────────────────────────────────────────────

    def open_file_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(self, tr("PDF openen"), "", tr("PDF-bestanden (*.pdf)"))
        if file_path:
            self.open_file(file_path)

    def open_file(self, file_path):
        try:
            tab = PdfTabWidget(file_path, parent=self.tabs)
        except Exception as exc:
            QMessageBox.critical(self, tr("Kan PDF niet openen"), f"{file_path}\n\n{exc}")
            return
        tab.set_thumbnails_visible(self._thumbnails_visible)
        tab.view.stateChanged.connect(self._update_actions_enabled)
        tab.view.undoAvailable.connect(lambda _available=None: self._update_actions_enabled())
        tab.view.pageChanged.connect(lambda page_num, t=tab: self._on_page_changed_for_tab(t, page_num))
        index = self.tabs.addTab(tab, tab.title)
        self.tabs.setCurrentIndex(index)
        self._update_actions_enabled()
        self._update_central_widget()
        settings.add_recent_file(file_path)

    # ── Printen & opslaan ────────────────────────────────────────────

    def _print_current(self):
        tab = self.tabs.currentWidget()
        if tab is None or not tab.view.pdf_document:
            return

        dialog = PrintDialog(self, tab.title, tab.page_count, tab.view.current_page)
        if not dialog.exec():
            return

        options = dialog.get_options()
        printer = dialog.printer
        printer.setDuplex(QPrinter.DuplexMode.DuplexLongSide if options.duplex else QPrinter.DuplexMode.DuplexNone)
        printer.setPageOrientation(
            QPageLayout.Orientation.Landscape if options.orientation == "liggend" else QPageLayout.Orientation.Portrait
        )

        total_units = len(options.pages) * max(options.copies, 1)
        progress = QProgressDialog(tr("Bezig met printen..."), tr("Annuleren"), 0, total_units, self)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)

        def on_progress(done, total):
            progress.setValue(done)
            progress.setLabelText(tr("Pagina {done} van {total}", done=done, total=total))

        try:
            completed = print_backend.run_print_job(
                printer, tab.view.pdf_document, options, on_progress, progress.wasCanceled
            )
        except Exception as exc:
            progress.close()
            QMessageBox.critical(self, tr("Printfout"), tr("Kan niet printen:") + f"\n\n{exc}")
            return

        progress.close()
        if not completed:
            QMessageBox.information(self, tr("Geannuleerd"), tr("Het printen is geannuleerd."))

    def _save_as_current(self):
        tab = self.tabs.currentWidget()
        if tab is not None:
            save_pdf.save_as(self, tab)

    def _send_current(self):
        tab = self.tabs.currentWidget()
        if tab is not None:
            email_sender.send_as_attachment(self, tab)

    # ── Help ─────────────────────────────────────────────────────────

    def _show_pdf_info(self):
        """Poort van NVict_Reader.py:5057-5073 (show_pdf_info)."""
        tab = self.tabs.currentWidget()
        if tab is None or not tab.view.pdf_document:
            return
        metadata = tab.view.pdf_document.metadata or {}
        rows = [
            (tr("Titel"), metadata.get("title")),
            (tr("Auteur"), metadata.get("author")),
            (tr("Onderwerp"), metadata.get("subject")),
            (tr("Trefwoorden"), metadata.get("keywords")),
            ("Creator", metadata.get("creator")),
            ("Producer", metadata.get("producer")),
            (tr("Gemaakt"), metadata.get("creationDate")),
            (tr("Gewijzigd"), metadata.get("modDate")),
            (tr("Pagina's"), tab.page_count),
            (tr("Bestandsgrootte"), f"{os.path.getsize(tab.file_path) / 1024:.1f} KB"),
        ]
        info_text = "\n".join(f"{label}: {value or 'N/A'}" for label, value in rows)
        QMessageBox.information(self, tr("PDF-informatie"), info_text)

    def _show_about(self):
        AboutDialog(self, SOFTWARE_URL).exec()

    # ── Instellingen ──────────────────────────────────────────────────

    def _open_settings(self):
        dialog = SettingsDialog(self)
        accepted = dialog.exec()
        # De recente-bestandenlijst kan in het venster al gewist zijn, ook bij Annuleren.
        self.welcome_widget.refresh()
        if not accepted:
            return
        mode = dialog.get_theme_mode()
        settings.save_theme_mode(mode)
        theme.apply_theme(QApplication.instance(), mode)
        self._refresh_icons()
        self.welcome_widget.refresh()

        show_thumbnails = dialog.get_show_thumbnails_default()
        settings.save_show_thumbnails_default(show_thumbnails)
        self.action_toggle_thumbnails.setChecked(show_thumbnails)

        language_changed = dialog.language_changed()
        settings.save_language(dialog.get_language())
        if language_changed:
            # Bewust in de NIEUWE taal: dat is de taal die de gebruiker net koos.
            if i18n.resolve(dialog.get_language()) == i18n.LANGUAGE_ENGLISH:
                title, text = "Language", "The new language will be used the next time you start NVict Reader."
            else:
                title, text = "Taal", "De nieuwe taal wordt gebruikt zodra u NVict Reader opnieuw start."
            QMessageBox.information(self, title, text)

    # ── Bewerken-menu: exporteren, samenvoegen, roteren ──────────────

    def _export_pages_current(self):
        tab = self.tabs.currentWidget()
        if tab is None or not tab.view.pdf_document:
            return
        if not save_pdf.confirm_discard_unsaved(self, tab.view):
            return

        dialog = ExportPagesDialog(self, tab.page_count)
        if not dialog.exec():
            return
        pages = dialog.get_pages()

        suggested = os.path.splitext(tab.file_path)[0] + "_export.pdf"
        target_path, _ = QFileDialog.getSaveFileName(self, tr("Pagina's exporteren"), suggested, tr("PDF-bestanden (*.pdf)"))
        if not target_path:
            return

        try:
            document_tools.export_pages(tab.view.pdf_document, pages, target_path)
        except Exception as exc:
            QMessageBox.critical(self, tr("Fout"), tr("Kan pagina's niet exporteren:") + f"\n{exc}")
            return
        QMessageBox.information(
            self, tr("Succes"),
            tr("{count} pagina('s) succesvol geëxporteerd naar:\n{path}", count=len(pages), path=target_path),
        )

    def _merge_pdfs_current(self):
        open_paths = [self.tabs.widget(i).file_path for i in range(self.tabs.count())]
        dialog = MergePdfsDialog(self, open_paths)
        if not dialog.exec():
            return
        file_paths = dialog.get_file_paths()

        target_path, _ = QFileDialog.getSaveFileName(
            self, tr("PDF's samenvoegen"), tr("samengevoegd.pdf"), tr("PDF-bestanden (*.pdf)")
        )
        if not target_path:
            return

        try:
            document_tools.merge_pdfs(file_paths, target_path)
        except Exception as exc:
            QMessageBox.critical(self, tr("Fout"), tr("Kan PDF's niet samenvoegen:") + f"\n{exc}")
            return

        reply = QMessageBox.question(
            self, tr("Succes"),
            tr("{count} bestanden succesvol samengevoegd naar:\n{path}\n\nNu openen?", count=len(file_paths), path=target_path),
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.open_file(target_path)

    def _rotate_pages_current(self):
        tab = self.tabs.currentWidget()
        if tab is None or not tab.view.pdf_document:
            return

        dialog = RotatePagesDialog(self, tab.page_count, tab.view.current_page)
        if not dialog.exec():
            return
        pages = dialog.get_pages()
        degrees = dialog.get_rotation()

        try:
            tab.view.rotate_pages(pages, degrees)
        except Exception as exc:
            QMessageBox.critical(self, tr("Fout"), tr("Kan pagina's niet roteren:") + f"\n{exc}")
            return
        QMessageBox.information(
            self, tr("Geroteerd"),
            tr("{count} pagina('s) geroteerd met {degrees}°.\n\nVergeet niet op te slaan om de wijziging te behouden!",
               count=len(pages), degrees=degrees),
        )

    # ── Window lifecycle ─────────────────────────────────────────────

    def closeEvent(self, event):
        for index in range(self.tabs.count()):
            widget = self.tabs.widget(index)
            if widget is not None and not save_pdf.confirm_discard_unsaved(self, widget.view):
                event.ignore()
                return

        settings.save_window_state(self)
        for index in range(self.tabs.count()):
            widget = self.tabs.widget(index)
            if widget is not None:
                widget.close_document()
        super().closeEvent(event)
