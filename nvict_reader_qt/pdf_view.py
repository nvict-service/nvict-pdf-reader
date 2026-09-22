# -*- coding: utf-8 -*-
"""PDF-render-widget: doorlopend scrollen, lazy render van zichtbare pagina's.

Qt-analogie van PDFTab/_render_visible_pages uit NVict_Reader.py: één
QGraphicsPixmapItem per zichtbare pagina op vaste scene-coördinaten, met
lichte placeholder-rects voor niet-zichtbare pagina's zodat de scrollbars
meteen de juiste documentgrootte kennen.
"""

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtGui import QBrush, QColor, QImage, QPainter, QPen, QPixmap
from PySide6.QtWidgets import QGraphicsPixmapItem, QGraphicsRectItem, QGraphicsScene, QGraphicsView

from . import rendering
from .document import get_fitz

RENDER_DEBOUNCE_MS = 60
PLACEHOLDER_COLOR = QColor("#d9d9d9")


class PdfGraphicsView(QGraphicsView):
    """Toont één PDF-document met doorlopend scrollen en lazy rendering."""

    zoomChanged = Signal(float)
    pageChanged = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setScene(QGraphicsScene(self))
        self.setRenderHint(QPainter.RenderHint.Antialiasing, False)
        self.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)

        self.pdf_document = None
        self.file_path = None
        self.zoom_level = 1.0
        self.zoom_mode = "fit_width"
        self.current_page = 0

        self.page_layout = []
        self._layout_by_page = {}
        self._pixmap_items = {}
        self._placeholder_items = {}
        self._pixmap_cache = {}
        self._rendered_pages = None

        self._render_timer = QTimer(self)
        self._render_timer.setSingleShot(True)
        self._render_timer.timeout.connect(self.render_visible_pages)

        self.verticalScrollBar().valueChanged.connect(self._schedule_render)
        self.horizontalScrollBar().valueChanged.connect(self._schedule_render)

    # ── Document lifecycle ────────────────────────────────────────────

    def open_document(self, file_path, password=None):
        """Open een PDF-bestand en toon pagina 1."""
        self.close_document()
        fitz = get_fitz()
        document = fitz.open(file_path)
        if password and document.needs_pass:
            document.authenticate(password)

        self.pdf_document = document
        self.file_path = file_path
        self.current_page = 0
        self.zoom_mode = "fit_width"
        self._recompute_fit_width_zoom()
        self.rebuild_layout()

    def close_document(self):
        if self.pdf_document is not None:
            try:
                self.pdf_document.close()
            except Exception:
                pass
        self.pdf_document = None
        self.scene().clear()
        self.page_layout = []
        self._layout_by_page = {}
        self._pixmap_items = {}
        self._placeholder_items = {}
        self._pixmap_cache = {}
        self._rendered_pages = None

    # ── Layout & rendering ───────────────────────────────────────────

    def rebuild_layout(self, force_render=True):
        """Herbereken pagina-posities en herbouw de scene (placeholders)."""
        if not self.pdf_document:
            return

        self.scene().clear()
        self._pixmap_items = {}
        self._placeholder_items = {}
        self._pixmap_cache = {}
        self._rendered_pages = None

        layout, total_width, total_height = rendering.compute_page_layout(self.pdf_document, self.zoom_level)
        self.page_layout = layout
        self._layout_by_page = {entry.page_num: entry for entry in layout}
        self.scene().setSceneRect(0, 0, total_width, total_height)

        for entry in layout:
            placeholder = QGraphicsRectItem(0, 0, entry.width, entry.height)
            placeholder.setPos(entry.x, entry.y)
            placeholder.setBrush(QBrush(PLACEHOLDER_COLOR))
            placeholder.setPen(QPen(Qt.PenStyle.NoPen))
            self.scene().addItem(placeholder)
            self._placeholder_items[entry.page_num] = placeholder

        if force_render:
            self.render_visible_pages(force=True)

    def _schedule_render(self, *_args):
        self._render_timer.start(RENDER_DEBOUNCE_MS)

    def render_visible_pages(self, force=False):
        """Render alleen de pagina's die in beeld staan (plus marge)."""
        if not self.pdf_document or not self.page_layout:
            return

        viewport_rect = self.mapToScene(self.viewport().rect()).boundingRect()
        wanted = rendering.visible_page_numbers(
            self.page_layout, viewport_rect.top(), viewport_rect.height(), self.current_page
        )
        if not force and wanted == self._rendered_pages:
            return

        wanted_set = set(wanted)

        for page_num in list(self._pixmap_items.keys()):
            if page_num not in wanted_set:
                item = self._pixmap_items.pop(page_num)
                self.scene().removeItem(item)
                self._restore_placeholder(page_num)

        for page_num in wanted:
            if page_num in self._pixmap_items:
                continue
            pixmap = self._pixmap_cache.get(page_num)
            if pixmap is None:
                pixmap = self._render_page_pixmap(page_num)
                self._pixmap_cache[page_num] = pixmap

            placeholder = self._placeholder_items.pop(page_num, None)
            if placeholder is not None:
                self.scene().removeItem(placeholder)

            entry = self._layout_by_page[page_num]
            item = QGraphicsPixmapItem(pixmap)
            item.setPos(entry.x, entry.y)
            item.setZValue(1)
            self.scene().addItem(item)
            self._pixmap_items[page_num] = item

        self._rendered_pages = wanted
        anchor = wanted[0] if wanted else self.current_page
        rendering.trim_page_cache(self._pixmap_cache, wanted, anchor)

        new_current = self.current_page_number()
        if new_current != self.current_page:
            self.current_page = new_current
            self.pageChanged.emit(new_current)

    def _restore_placeholder(self, page_num):
        if page_num in self._placeholder_items:
            return
        entry = self._layout_by_page.get(page_num)
        if entry is None:
            return
        placeholder = QGraphicsRectItem(0, 0, entry.width, entry.height)
        placeholder.setPos(entry.x, entry.y)
        placeholder.setBrush(QBrush(PLACEHOLDER_COLOR))
        placeholder.setPen(QPen(Qt.PenStyle.NoPen))
        self.scene().addItem(placeholder)
        self._placeholder_items[page_num] = placeholder

    def _render_page_pixmap(self, page_num) -> QPixmap:
        fitz = get_fitz()
        page = self.pdf_document[page_num]
        mat = fitz.Matrix(self.zoom_level, self.zoom_level)
        pix = page.get_pixmap(matrix=mat)
        fmt = QImage.Format.Format_RGBA8888 if pix.alpha else QImage.Format.Format_RGB888
        image = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
        return QPixmap.fromImage(image.copy())

    # ── Zoom & navigatie ─────────────────────────────────────────────

    def _recompute_fit_width_zoom(self):
        viewport_width = self.viewport().width()
        if viewport_width <= 1:
            viewport_width = 800
        self.zoom_level = rendering.fit_width_zoom(self.pdf_document, viewport_width)

    def set_zoom_mode_fit_width(self):
        if not self.pdf_document:
            return
        self.zoom_mode = "fit_width"
        self._recompute_fit_width_zoom()
        self.rebuild_layout()
        self.zoomChanged.emit(self.zoom_level)

    def zoom_in(self):
        self._set_manual_zoom(self.zoom_level * 1.25)

    def zoom_out(self):
        self._set_manual_zoom(self.zoom_level * 0.8)

    def _set_manual_zoom(self, new_zoom):
        if not self.pdf_document:
            return
        self.zoom_mode = "manual"
        self.zoom_level = max(0.1, min(new_zoom, 8.0))
        self.rebuild_layout()
        self.zoomChanged.emit(self.zoom_level)

    def current_page_number(self):
        if not self.page_layout:
            return 0
        viewport_rect = self.mapToScene(self.viewport().rect()).boundingRect()
        top = viewport_rect.top()
        for entry in self.page_layout:
            if entry.y + entry.height > top:
                return entry.page_num
        return self.page_layout[-1].page_num

    def go_to_page(self, page_num):
        if not self.page_layout:
            return
        page_num = max(0, min(page_num, len(self.page_layout) - 1))
        entry = self._layout_by_page[page_num]
        self.centerOn(entry.x + entry.width / 2, entry.y + self.viewport().height() / 3)
        self.current_page = page_num
        self.pageChanged.emit(page_num)
        self.render_visible_pages(force=True)

    def next_page(self):
        self.go_to_page(self.current_page + 1)

    def prev_page(self):
        self.go_to_page(self.current_page - 1)

    def first_page(self):
        self.go_to_page(0)

    def last_page(self):
        if self.page_layout:
            self.go_to_page(len(self.page_layout) - 1)

    # ── Qt events ────────────────────────────────────────────────────

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.pdf_document and self.zoom_mode == "fit_width":
            self._recompute_fit_width_zoom()
            self.rebuild_layout()
