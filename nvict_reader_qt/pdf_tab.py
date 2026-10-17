# -*- coding: utf-8 -*-
"""Eén tabblad: een geopend PDF-document met eigen scroll/zoom-state.

Qt-equivalent van PDFTab(tk.Frame) uit NVict_Reader.py (regel 547) - hier
een dunne wrapper om PdfGraphicsView + de (optionele) zijpanelen
(inhoudsopgave, paginaminiaturen) en de zoekbalk, zodat MainWindow per tab
kan delegeren zonder zelf documentstate te houden.
"""

import os

from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import QSplitter, QVBoxLayout, QWidget

from .pdf_view import PdfGraphicsView
from .search_bar import SearchBar
from .thumbnail_panel import ThumbnailPanel
from .toc_panel import TocPanel


class PdfTabWidget(QWidget):
    def __init__(self, file_path, password=None, parent=None):
        super().__init__(parent)
        self.view = PdfGraphicsView(self)
        self.toc_panel = TocPanel(self.view, self)
        self.toc_panel.setVisible(False)
        self.thumbnail_panel = ThumbnailPanel(self.view, self)
        self.thumbnail_panel.setVisible(False)
        self._toc_wanted = False

        self.splitter = QSplitter(self)
        self.splitter.addWidget(self.toc_panel)
        self.splitter.addWidget(self.thumbnail_panel)
        self.splitter.addWidget(self.view)
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 0)
        self.splitter.setStretchFactor(2, 1)
        self.splitter.setSizes([220, 160, 1000])

        self.search_bar = SearchBar(self.view, self)
        self.search_bar.setVisible(False)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(self.splitter, 1)
        layout.addWidget(self.search_bar)

        QShortcut(QKeySequence(QKeySequence.StandardKey.FindNext), self, self.search_bar.find_next)
        QShortcut(QKeySequence(QKeySequence.StandardKey.FindPrevious), self, self.search_bar.find_previous)

        self.view.open_document(file_path, password=password)
        self.view.structureChanged.connect(self.refresh_panels)
        self.refresh_panels()

    def refresh_panels(self):
        """Na openen of na een paginabewerking (verwijderen/verplaatsen/invoegen)."""
        self.thumbnail_panel.rebuild()
        self.toc_panel.rebuild()
        self._apply_toc_visibility()

    @property
    def file_path(self):
        return self.view.file_path

    @property
    def title(self):
        return os.path.basename(self.file_path) if self.file_path else "PDF"

    @property
    def page_count(self):
        return len(self.view.pdf_document) if self.view.pdf_document else 0

    def has_toc(self) -> bool:
        return self.toc_panel.has_entries()

    def set_thumbnails_visible(self, visible: bool):
        self.thumbnail_panel.setVisible(visible)

    def set_toc_visible(self, visible: bool):
        self._toc_wanted = visible
        self._apply_toc_visibility()

    def _apply_toc_visibility(self):
        self.toc_panel.setVisible(self._toc_wanted and self.has_toc())

    def show_search_bar(self):
        self.search_bar.open_bar()

    def close_document(self):
        self.view.close_document()
