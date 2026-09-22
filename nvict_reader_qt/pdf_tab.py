# -*- coding: utf-8 -*-
"""Eén tabblad: een geopend PDF-document met eigen scroll/zoom-state.

Qt-equivalent van PDFTab(tk.Frame) uit NVict_Reader.py (regel 547) - hier
een dunne wrapper om PdfGraphicsView, zodat MainWindow per tab kan
delegeren zonder zelf documentstate te houden.
"""

import os

from PySide6.QtWidgets import QVBoxLayout, QWidget

from .pdf_view import PdfGraphicsView


class PdfTabWidget(QWidget):
    def __init__(self, file_path, password=None, parent=None):
        super().__init__(parent)
        self.view = PdfGraphicsView(self)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.view)

        self.view.open_document(file_path, password=password)

    @property
    def file_path(self):
        return self.view.file_path

    @property
    def title(self):
        return os.path.basename(self.file_path) if self.file_path else "PDF"

    @property
    def page_count(self):
        return len(self.view.pdf_document) if self.view.pdf_document else 0

    def close_document(self):
        self.view.close_document()
