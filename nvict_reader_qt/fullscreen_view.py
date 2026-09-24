# -*- coding: utf-8 -*-
"""Volledig-scherm presentatiemodus voor de huidige pagina.

Poort van NVict_Reader.py:1429-1580 (enter_fullscreen/_render_fs_page/
exit_fullscreen): een apart, randloos venster dat de huidige pagina
letterboxed op zwart toont, met een tijdelijke hint-banner en
pijltjes/PageUp/PageDown-navigatie los van de hoofd-tab. Bij sluiten wordt
de hoofd-tab gesynchroniseerd naar de laatst bekeken pagina.
"""

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QImage, QPixmap
from PySide6.QtWidgets import QLabel, QVBoxLayout, QWidget

from .document import get_fitz
from .i18n import tr

HINT_HIDE_DELAY_MS = 4000
HINT_TEXT = "Escape / F11 om te verlaten   ·   ← → voor navigatie"


class FullscreenWindow(QWidget):
    def __init__(self, pdf_document, start_page, on_close):
        super().__init__()
        self.pdf_document = pdf_document
        self.page_num = start_page
        self._on_close = on_close

        self.setWindowFlag(Qt.WindowType.Window, True)
        self.setStyleSheet("background-color: black;")
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.page_label = QLabel(self)
        self.page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.page_label)

        self.hint_label = QLabel(tr(HINT_TEXT), self)
        self.hint_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.hint_label.setStyleSheet(
            "background-color: #1e1e1e; color: #dddddd; font-size: 14px; padding: 14px;"
        )

        self.info_label = QLabel(self)
        self.info_label.setStyleSheet("background-color: black; color: #666666; font-size: 10px;")
        self.info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self._hint_timer = QTimer(self)
        self._hint_timer.setSingleShot(True)
        self._hint_timer.timeout.connect(self.hint_label.hide)

        self.showFullScreen()
        self._render_page()
        self._show_hint()

    def _show_hint(self):
        self.hint_label.setGeometry(0, 0, self.width(), self.hint_label.sizeHint().height())
        self.hint_label.show()
        self.hint_label.raise_()
        self._hint_timer.start(HINT_HIDE_DELAY_MS)

    def _render_page(self):
        if not self.pdf_document:
            return
        screen_w, screen_h = self.width(), self.height()
        if screen_w < 10 or screen_h < 10:
            return

        fitz = get_fitz()
        page = self.pdf_document[self.page_num]
        bound = page.bound()
        zoom = min(screen_w / bound.width, screen_h / bound.height) if bound.width and bound.height else 1.0

        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        fmt = QImage.Format.Format_RGBA8888 if pix.alpha else QImage.Format.Format_RGB888
        image = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
        self.page_label.setPixmap(QPixmap.fromImage(image.copy()))

        total = len(self.pdf_document)
        self.info_label.setText(tr("Pagina {page} / {total}  ·  Escape om te sluiten", page=self.page_num + 1, total=total))
        self.info_label.setGeometry(
            0, screen_h - self.info_label.sizeHint().height() - 10, screen_w, self.info_label.sizeHint().height()
        )
        self.info_label.raise_()

    def _navigate(self, delta):
        if not self.pdf_document:
            return
        new_page = self.page_num + delta
        if 0 <= new_page < len(self.pdf_document):
            self.page_num = new_page
            self._render_page()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._render_page()
        if self.hint_label.isVisible():
            self.hint_label.setGeometry(0, 0, self.width(), self.hint_label.sizeHint().height())

    def mouseMoveEvent(self, event):
        super().mouseMoveEvent(event)
        if event.position().y() < 80 and not self.hint_label.isVisible():
            self._show_hint()

    def keyPressEvent(self, event):
        key = event.key()
        if key in (Qt.Key.Key_Escape, Qt.Key.Key_F11):
            self.close()
        elif key in (Qt.Key.Key_Left, Qt.Key.Key_PageUp):
            self._navigate(-1)
        elif key in (Qt.Key.Key_Right, Qt.Key.Key_PageDown):
            self._navigate(1)
        else:
            super().keyPressEvent(event)

    def closeEvent(self, event):
        if self._on_close is not None:
            self._on_close(self.page_num)
        super().closeEvent(event)
