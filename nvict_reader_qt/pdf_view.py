# -*- coding: utf-8 -*-
"""PDF-render-widget: doorlopend scrollen, lazy render van zichtbare pagina's.

Qt-analogie van PDFTab/_render_visible_pages uit NVict_Reader.py: één
QGraphicsPixmapItem per zichtbare pagina op vaste scene-coördinaten, met
lichte placeholder-rects voor niet-zichtbare pagina's zodat de scrollbars
meteen de juiste documentgrootte kennen.
"""

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtGui import QBrush, QColor, QFont, QImage, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QGraphicsPixmapItem,
    QGraphicsRectItem,
    QGraphicsScene,
    QGraphicsSimpleTextItem,
    QGraphicsView,
    QMenu,
)

from . import rendering
from .annotations import COLOR_MAP, HIGHLIGHT_COLOR, HighlightAnnotation, TextAnnotation
from .document import get_fitz
from .text_annotation_dialog import TextAnnotationDialog

RENDER_DEBOUNCE_MS = 60
PLACEHOLDER_COLOR = QColor("#d9d9d9")
DRAG_SELECTION_BRUSH = QBrush(QColor(173, 216, 230, 100))
DRAG_SELECTION_PEN = QPen(QColor(100, 150, 200, 200))

_FONT_FAMILY_BY_CODE = {"helv": "Arial", "tiro": "Times New Roman", "cour": "Courier New"}


def _qcolor_from_pdf(color_key) -> QColor:
    r, g, b = COLOR_MAP.get(color_key, (0, 0, 0))
    return QColor(int(r * 255), int(g * 255), int(b * 255))


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

        # ── Annotatie/markeer-state (fase 2) ──
        self.tool_mode = None  # None / "text_annotate" / "highlight"
        self.text_annotations = []       # list[TextAnnotation]
        self.highlight_annotations = []  # list[HighlightAnnotation]
        self._text_overlay_items = []    # list[(TextAnnotation, QGraphicsSimpleTextItem)]
        self._word_cache = {}            # page_num -> list[WordBox], lazy per gebruik
        self._drag_rect_item = None
        self._drag_start_scene = None

    def has_unsaved_changes(self):
        return bool(self.text_annotations) or bool(self.highlight_annotations)

    def clear_saved_annotations(self):
        """Na succesvol 'Opslaan als': wis de pending-state.

        Tekst-annotaties waren nog geen echte PDF-annotatie, dus hun overlay
        moet ook van het scherm; highlights stonden al op het live
        fitz-document en blijven dus gewoon zichtbaar (alleen de
        boekhoudlijst wordt gewist, zoals in NVict_Reader.py:4179-4183).
        """
        for _annotation, item in self._text_overlay_items:
            self.scene().removeItem(item)
        self._text_overlay_items = []
        self.text_annotations = []
        self.highlight_annotations = []

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
        self.text_annotations = []
        self.highlight_annotations = []
        self._text_overlay_items = []
        self._word_cache = {}
        self._drag_rect_item = None
        self._drag_start_scene = None

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
        self._text_overlay_items = []
        self._drag_rect_item = None

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

        for annotation in self.text_annotations:
            self._add_text_annotation_overlay(annotation)

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

    # ── Annotaties: tekst & markeren (fase 2) ──────────────────────────

    def set_tool_mode(self, mode):
        """Zet de actieve tool ("text_annotate"/"highlight"/None), exclusief."""
        self.tool_mode = mode
        if mode == "text_annotate":
            self.setCursor(Qt.CursorShape.IBeamCursor)
        elif mode == "highlight":
            self.setCursor(Qt.CursorShape.CrossCursor)
        else:
            self.unsetCursor()

    def _page_at_scene_pos(self, pos):
        """Geef de PageLayoutEntry terug waar `pos` (scene-coördinaten) binnen valt."""
        for entry in self.page_layout:
            if entry.x <= pos.x() <= entry.x + entry.width and entry.y <= pos.y() <= entry.y + entry.height:
                return entry
        return None

    def _words_for_page(self, page_num):
        words = self._word_cache.get(page_num)
        if words is None:
            words = rendering.get_page_words(self.pdf_document, page_num)
            self._word_cache[page_num] = words
            rendering.trim_page_cache(self._word_cache, keep=[page_num], anchor=self.current_page,
                                       max_pages=rendering.WORD_CACHE_PAGES)
        return words

    # ── Tekst-annotaties ──

    def _add_text_annotation_overlay(self, annotation: TextAnnotation):
        entry = self._layout_by_page.get(annotation.page_num)
        if entry is None:
            return
        item = QGraphicsSimpleTextItem(annotation.text)
        font = QFont(_FONT_FAMILY_BY_CODE.get(annotation.fontname, "Arial"))
        font.setPixelSize(max(1, round(annotation.font_size * self.zoom_level)))
        item.setFont(font)
        item.setBrush(QBrush(_qcolor_from_pdf(annotation.color)))
        item.setPos(entry.x + annotation.pdf_x * self.zoom_level, entry.y + annotation.pdf_y * self.zoom_level)
        item.setZValue(2)
        self.scene().addItem(item)
        self._text_overlay_items.append((annotation, item))

    def _text_annotation_at(self, pos):
        for annotation, item in self._text_overlay_items:
            if item.contains(item.mapFromScene(pos)):
                return annotation, item
        return None

    def _handle_text_annotate_click(self, scene_pos):
        hit = self._text_annotation_at(scene_pos)
        if hit is not None:
            annotation, item = hit
            existing = {
                "text": annotation.text, "font_size": annotation.font_size,
                "color": annotation.color, "fontname": annotation.fontname,
            }
            dialog = TextAnnotationDialog(self, existing=existing)
            if dialog.exec():
                if dialog.delete_requested():
                    self.text_annotations.remove(annotation)
                    self._text_overlay_items.remove((annotation, item))
                    self.scene().removeItem(item)
                else:
                    result = dialog.get_result()
                    if result["text"]:
                        annotation.text = result["text"]
                        annotation.font_size = result["font_size"]
                        annotation.color = result["color"]
                        annotation.fontname = result["fontname"]
                        self._text_overlay_items.remove((annotation, item))
                        self.scene().removeItem(item)
                        self._add_text_annotation_overlay(annotation)
            return

        entry = self._page_at_scene_pos(scene_pos)
        if entry is None:
            return
        pdf_x = (scene_pos.x() - entry.x) / self.zoom_level
        pdf_y_click = (scene_pos.y() - entry.y) / self.zoom_level

        dialog = TextAnnotationDialog(self)
        if dialog.exec():
            result = dialog.get_result()
            if result["text"]:
                # Y-correctie zodat het klikpunt de onderkant van de tekst is,
                # zelfde correctie als NVict_Reader.py:4722.
                annotation = TextAnnotation(
                    page_num=entry.page_num, pdf_x=pdf_x, pdf_y=pdf_y_click - result["font_size"],
                    text=result["text"], font_size=result["font_size"],
                    color=result["color"], fontname=result["fontname"],
                )
                self.text_annotations.append(annotation)
                self._add_text_annotation_overlay(annotation)

    # ── Markeren (highlight) ──

    def _handle_highlight_press(self, scene_pos):
        self._drag_start_scene = scene_pos
        self._drag_rect_item = QGraphicsRectItem(scene_pos.x(), scene_pos.y(), 0, 0)
        self._drag_rect_item.setBrush(DRAG_SELECTION_BRUSH)
        self._drag_rect_item.setPen(DRAG_SELECTION_PEN)
        self._drag_rect_item.setZValue(3)
        self.scene().addItem(self._drag_rect_item)

    def _handle_highlight_move(self, scene_pos):
        if self._drag_rect_item is None or self._drag_start_scene is None:
            return
        x1, y1 = self._drag_start_scene.x(), self._drag_start_scene.y()
        x2, y2 = scene_pos.x(), scene_pos.y()
        rect_x, rect_y = min(x1, x2), min(y1, y2)
        self._drag_rect_item.setRect(rect_x, rect_y, abs(x2 - x1), abs(y2 - y1))

    def _handle_highlight_release(self, scene_pos):
        if self._drag_rect_item is not None:
            self.scene().removeItem(self._drag_rect_item)
            self._drag_rect_item = None
        if self._drag_start_scene is None:
            return

        left = min(self._drag_start_scene.x(), scene_pos.x())
        right = max(self._drag_start_scene.x(), scene_pos.x())
        top = min(self._drag_start_scene.y(), scene_pos.y())
        bottom = max(self._drag_start_scene.y(), scene_pos.y())
        self._drag_start_scene = None

        fitz = get_fitz()
        for entry in self.page_layout:
            page_right = entry.x + entry.width
            page_bottom = entry.y + entry.height
            if entry.x > right or page_right < left or entry.y > bottom or page_bottom < top:
                continue

            quads = []
            for word in self._words_for_page(entry.page_num):
                wx0 = entry.x + word.x0 * self.zoom_level
                wy0 = entry.y + word.y0 * self.zoom_level
                wx1 = entry.x + word.x1 * self.zoom_level
                wy1 = entry.y + word.y1 * self.zoom_level
                if wx1 < left or wx0 > right or wy1 < top or wy0 > bottom:
                    continue
                quads.append(fitz.Rect(word.x0, word.y0, word.x1, word.y1).quad)

            if quads:
                self._apply_highlight(entry.page_num, quads)

    def _apply_highlight(self, page_num, quads):
        page = self.pdf_document[page_num]
        annot = page.add_highlight_annot(quads)
        annot.set_colors(stroke=HIGHLIGHT_COLOR)
        annot.update()
        self.highlight_annotations.append(HighlightAnnotation(page_num=page_num, quads=quads, xref=annot.xref))
        self._pixmap_cache.pop(page_num, None)
        self.render_visible_pages(force=True)

    def _highlight_at(self, entry, scene_pos):
        pdf_x = (scene_pos.x() - entry.x) / self.zoom_level
        pdf_y = (scene_pos.y() - entry.y) / self.zoom_level
        for highlight in self.highlight_annotations:
            if highlight.page_num != entry.page_num:
                continue
            for quad in highlight.quads:
                rect = quad.rect
                if rect.x0 <= pdf_x <= rect.x1 and rect.y0 <= pdf_y <= rect.y1:
                    return highlight
        return None

    def _remove_highlight(self, highlight):
        page = self.pdf_document[highlight.page_num]
        try:
            annot = page.load_annot(highlight.xref)
            page.delete_annot(annot)
        except Exception:
            pass
        if highlight in self.highlight_annotations:
            self.highlight_annotations.remove(highlight)
        self._pixmap_cache.pop(highlight.page_num, None)
        self.render_visible_pages(force=True)

    def _handle_right_click(self, scene_pos):
        entry = self._page_at_scene_pos(scene_pos)
        if entry is None:
            return
        highlight = self._highlight_at(entry, scene_pos)
        if highlight is None:
            return
        menu = QMenu(self)
        remove_action = menu.addAction("Markering verwijderen")
        chosen = menu.exec(self.viewport().mapToGlobal(self.mapFromScene(scene_pos)))
        if chosen == remove_action:
            self._remove_highlight(highlight)

    # ── Qt events ────────────────────────────────────────────────────

    def mousePressEvent(self, event):
        if not self.pdf_document:
            super().mousePressEvent(event)
            return

        scene_pos = self.mapToScene(event.position().toPoint())

        if event.button() == Qt.MouseButton.RightButton:
            self._handle_right_click(scene_pos)
            return

        if event.button() == Qt.MouseButton.LeftButton:
            if self.tool_mode == "text_annotate":
                self._handle_text_annotate_click(scene_pos)
                return
            if self.tool_mode == "highlight":
                self._handle_highlight_press(scene_pos)
                return

        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.tool_mode == "highlight" and self._drag_rect_item is not None:
            self._handle_highlight_move(self.mapToScene(event.position().toPoint()))
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self.tool_mode == "highlight" and self._drag_rect_item is not None:
            self._handle_highlight_release(self.mapToScene(event.position().toPoint()))
            return
        super().mouseReleaseEvent(event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.pdf_document and self.zoom_mode == "fit_width":
            self._recompute_fit_width_zoom()
            self.rebuild_layout()
