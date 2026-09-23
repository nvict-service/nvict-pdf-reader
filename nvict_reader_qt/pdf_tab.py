# -*- coding: utf-8 -*-
"""Eén tabblad: een geopend PDF-document met eigen scroll/zoom-state.

Qt-equivalent van PDFTab(tk.Frame) uit NVict_Reader.py (regel 547) - hier
een dunne wrapper om PdfGraphicsView + het (optionele) thumbnails-paneel,
zodat MainWindow per tab kan delegeren zonder zelf documentstate te houden.
"""

import os

from PySide6.QtWidgets import QHBoxLayout, QSplitter, QWidget

from .pdf_view import PdfGraphicsView
from .thumbnail_panel import ThumbnailPanel


class PdfTabWidget(QWidget):
    def __init__(self, file_path, password=None, parent=None):
        super().__init__(parent)
        self.view = PdfGraphicsView(self)
        self.thumbnail_panel = ThumbnailPanel(self.view, self)
        self.thumbnail_panel.setVisible(False)

        self.splitter = QSplitter(self)
        self.splitter.addWidget(self.thumbnail_panel)
        self.splitter.addWidget(self.view)
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 1)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.splitter)

        self.view.open_document(file_path, password=password)
        self.thumbnail_panel.rebuild()

    @property
    def file_path(self):
        return self.view.file_path

    @property
    def title(self):
        return os.path.basename(self.file_path) if self.file_path else "PDF"

    @property
    def page_count(self):
        return len(self.view.pdf_document) if self.view.pdf_document else 0

    def set_thumbnails_visible(self, visible: bool):
        self.thumbnail_panel.setVisible(visible)

    def close_document(self):
        self.view.close_document()
