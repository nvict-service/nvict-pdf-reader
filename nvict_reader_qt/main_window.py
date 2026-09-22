# -*- coding: utf-8 -*-
"""Hoofdvenster: menu, toolbar en tabbladen voor meerdere open PDF's.

Bevat zelf geen documentstate - delegeert naar de actieve tab
(PdfTabWidget), analoog aan hoe NVictReader in de tkinter-versie delegeert
naar self.get_active_tab().
"""

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QIcon, QKeySequence
from PySide6.QtWidgets import (
    QFileDialog,
    QMainWindow,
    QMessageBox,
    QTabWidget,
)

from . import settings
from .pdf_tab import PdfTabWidget

ICONS_DIR = Path(__file__).resolve().parent.parent / "icons"


def _icon(name: str) -> QIcon:
    path = ICONS_DIR / name
    return QIcon(str(path)) if path.exists() else QIcon()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("NVict Reader")
        self.resize(1366, 768)
        self.setMinimumSize(800, 600)

        self.tabs = QTabWidget(self)
        self.tabs.setTabsClosable(True)
        self.tabs.setMovable(True)
        self.tabs.tabCloseRequested.connect(self._close_tab)
        self.tabs.currentChanged.connect(self._on_tab_changed)
        self.setCentralWidget(self.tabs)

        self._build_actions()
        self._build_menu()
        self._build_toolbar()
        self._update_actions_enabled()

        settings.restore_window_state(self)

    # ── Actions ──────────────────────────────────────────────────────

    def _build_actions(self):
        self.action_open = QAction(_icon("open.png"), "&Openen...", self)
        self.action_open.setShortcut(QKeySequence.StandardKey.Open)
        self.action_open.triggered.connect(self.open_file_dialog)

        self.action_close_tab = QAction("Tab &sluiten", self)
        self.action_close_tab.setShortcut("Ctrl+W")
        self.action_close_tab.triggered.connect(lambda: self._close_tab(self.tabs.currentIndex()))

        self.action_quit = QAction("&Afsluiten", self)
        self.action_quit.setShortcut(QKeySequence.StandardKey.Quit)
        self.action_quit.triggered.connect(self.close)

        self.action_zoom_in = QAction(_icon("zoom-in.png"), "Zoom &in", self)
        self.action_zoom_in.setShortcut("Ctrl+=")
        self.action_zoom_in.triggered.connect(lambda: self._on_active(lambda v: v.zoom_in()))

        self.action_zoom_out = QAction(_icon("zoom-out.png"), "Zoom &uit", self)
        self.action_zoom_out.setShortcut("Ctrl+-")
        self.action_zoom_out.triggered.connect(lambda: self._on_active(lambda v: v.zoom_out()))

        self.action_fit_width = QAction(_icon("fit-width.png"), "&Pasbreedte", self)
        self.action_fit_width.setShortcut("Ctrl+0")
        self.action_fit_width.triggered.connect(lambda: self._on_active(lambda v: v.set_zoom_mode_fit_width()))

        self.action_first_page = QAction(_icon("first-page.png"), "&Eerste pagina", self)
        self.action_first_page.setShortcut("Ctrl+Home")
        self.action_first_page.triggered.connect(lambda: self._on_active(lambda v: v.first_page()))

        self.action_prev_page = QAction(_icon("prev-page.png"), "&Vorige pagina", self)
        self.action_prev_page.setShortcut(QKeySequence.StandardKey.MoveToPreviousPage)
        self.action_prev_page.triggered.connect(lambda: self._on_active(lambda v: v.prev_page()))

        self.action_next_page = QAction(_icon("next-page.png"), "&Volgende pagina", self)
        self.action_next_page.setShortcut(QKeySequence.StandardKey.MoveToNextPage)
        self.action_next_page.triggered.connect(lambda: self._on_active(lambda v: v.next_page()))

        self.action_last_page = QAction(_icon("last-page.png"), "&Laatste pagina", self)
        self.action_last_page.setShortcut("Ctrl+End")
        self.action_last_page.triggered.connect(lambda: self._on_active(lambda v: v.last_page()))

    def _build_menu(self):
        file_menu = self.menuBar().addMenu("&Bestand")
        file_menu.addAction(self.action_open)
        file_menu.addAction(self.action_close_tab)
        file_menu.addSeparator()
        file_menu.addAction(self.action_quit)

        view_menu = self.menuBar().addMenu("&Beeld")
        view_menu.addAction(self.action_zoom_in)
        view_menu.addAction(self.action_zoom_out)
        view_menu.addAction(self.action_fit_width)
        view_menu.addSeparator()
        view_menu.addAction(self.action_first_page)
        view_menu.addAction(self.action_prev_page)
        view_menu.addAction(self.action_next_page)
        view_menu.addAction(self.action_last_page)

    def _build_toolbar(self):
        toolbar = self.addToolBar("Hoofdwerkbalk")
        toolbar.setMovable(False)
        toolbar.addAction(self.action_open)
        toolbar.addSeparator()
        toolbar.addAction(self.action_first_page)
        toolbar.addAction(self.action_prev_page)
        toolbar.addAction(self.action_next_page)
        toolbar.addAction(self.action_last_page)
        toolbar.addSeparator()
        toolbar.addAction(self.action_zoom_out)
        toolbar.addAction(self.action_zoom_in)
        toolbar.addAction(self.action_fit_width)

    def _update_actions_enabled(self):
        has_tab = self.tabs.count() > 0
        for action in (
            self.action_close_tab, self.action_zoom_in, self.action_zoom_out,
            self.action_fit_width, self.action_first_page, self.action_prev_page,
            self.action_next_page, self.action_last_page,
        ):
            action.setEnabled(has_tab)

    # ── Tab helpers ──────────────────────────────────────────────────

    def _on_active(self, func):
        tab = self.tabs.currentWidget()
        if tab is not None:
            func(tab.view)

    def _on_tab_changed(self, _index):
        self._update_actions_enabled()

    def _close_tab(self, index):
        if index < 0:
            return
        tab = self.tabs.widget(index)
        self.tabs.removeTab(index)
        if tab is not None:
            tab.close_document()
            tab.deleteLater()
        self._update_actions_enabled()

    # ── File handling ────────────────────────────────────────────────

    def open_file_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "PDF openen", "", "PDF-bestanden (*.pdf)")
        if file_path:
            self.open_file(file_path)

    def open_file(self, file_path):
        try:
            tab = PdfTabWidget(file_path, parent=self.tabs)
        except Exception as exc:
            QMessageBox.critical(self, "Kan PDF niet openen", f"{file_path}\n\n{exc}")
            return
        index = self.tabs.addTab(tab, tab.title)
        self.tabs.setCurrentIndex(index)
        self._update_actions_enabled()

    # ── Window lifecycle ─────────────────────────────────────────────

    def closeEvent(self, event):
        settings.save_window_state(self)
        for index in range(self.tabs.count()):
            widget = self.tabs.widget(index)
            if widget is not None:
                widget.close_document()
        super().closeEvent(event)
