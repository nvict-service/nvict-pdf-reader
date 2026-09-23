# -*- coding: utf-8 -*-
"""PDF-render-widget: doorlopend scrollen, lazy render van zichtbare pagina's.

Qt-analogie van PDFTab/_render_visible_pages uit NVict_Reader.py: één
QGraphicsPixmapItem per zichtbare pagina op vaste scene-coördinaten, met
lichte placeholder-rects voor niet-zichtbare pagina's zodat de scrollbars
meteen de juiste documentgrootte kennen.
"""

import webbrowser

from PySide6.QtCore import QBuffer, QIODevice, Qt, QTimer, Signal
from PySide6.QtGui import QBrush, QColor, QFont, QImage, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QGraphicsItem,
    QGraphicsPixmapItem,
    QGraphicsRectItem,
    QGraphicsScene,
    QGraphicsSimpleTextItem,
    QGraphicsView,
    QMenu,
    QMessageBox,
)

from . import form_overlay, rendering, security
from .annotations import COLOR_MAP, HIGHLIGHT_COLOR, HighlightAnnotation, SignatureAnnotation, TextAnnotation
from .document import get_fitz
from .signature_dialog import SignatureDialog
from .text_annotation_dialog import TextAnnotationDialog

RENDER_DEBOUNCE_MS = 60
PLACEHOLDER_COLOR = QColor("#d9d9d9")
DRAG_SELECTION_BRUSH = QBrush(QColor(173, 216, 230, 100))
DRAG_SELECTION_PEN = QPen(QColor(100, 150, 200, 200))
FIELD_HIGHLIGHT_BRUSH = QBrush(QColor(208, 232, 255, 110))
FIELD_HIGHLIGHT_PEN = QPen(QColor(128, 184, 255))
FORM_WIDGETS_CACHE_PAGES = 20
LINKS_CACHE_PAGES = 20

_FONT_FAMILY_BY_CODE = {"helv": "Arial", "tiro": "Times New Roman", "cour": "Courier New"}


def _qcolor_from_pdf(color_key) -> QColor:
    r, g, b = COLOR_MAP.get(color_key, (0, 0, 0))
    return QColor(int(r * 255), int(g * 255), int(b * 255))


def _qimage_to_png_bytes(image: QImage) -> bytes:
    buffer = QBuffer()
    buffer.open(QIODevice.OpenModeFlag.WriteOnly)
    image.save(buffer, "PNG")
    return bytes(buffer.data())


class _SignatureItem(QGraphicsPixmapItem):
    """Versleepbare handtekening-afbeelding; houdt annotation.pdf_x/pdf_y bij.

    `ItemIsMovable` laat Qt het slepen zelf afhandelen (geen custom
    drag-code nodig); `itemChange` synchroniseert de PDF-coördinaten van de
    annotatie zodra de gebruiker de afbeelding verplaatst, zodat een
    her-render (bv. na zoomen) de nieuwe positie gebruikt in plaats van de
    oorspronkelijke plaatsingslocatie.
    """

    def __init__(self, pixmap, annotation, view):
        super().__init__(pixmap)
        self.annotation = annotation
        self._view = view
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, True)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges, True)
        # Getekende handtekeningen hebben een transparante achtergrond tussen
        # de penseelstreken door - zonder dit zou hit-testing (klikken om te
        # verslepen/verwijderen) alleen op de inkt zelf reageren, niet op de
        # hele omkaderde afbeelding zoals de gebruiker die ziet.
        self.setShapeMode(QGraphicsPixmapItem.ShapeMode.BoundingRectShape)

    def itemChange(self, change, value):
        if change == QGraphicsItem.GraphicsItemChange.ItemPositionHasChanged:
            entry = self._view._layout_by_page.get(self.annotation.page_num)
            if entry is not None and self._view.zoom_level:
                self.annotation.pdf_x = (value.x() - entry.x) / self._view.zoom_level
                self.annotation.pdf_y = (value.y() - entry.y) / self._view.zoom_level
        return super().itemChange(change, value)


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
        # Nodig zodat mouseMoveEvent ook zonder ingedrukte knop binnenkomt,
        # voor de link-hovercursor (fase 6).
        self.setMouseTracking(True)

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

        # ── Bewerken-menu-state (fase 3) ──
        self.pending_rotations = {}  # page_num -> graden, nog niet naar schijf geschreven

        # ── Formuliervelden (fase 4) ──
        self.form_mode = False           # apart van tool_mode: geen klik-plaats-tool, aanhoudende weergave
        self.form_field_values = {}      # xref -> waarde (bool voor checkbox/radio, str voor tekst/combobox)
        self._form_overlay_items = []    # list[QGraphicsProxyWidget], alleen actief tijdens form_mode
        self._radio_groups = {}          # field_name -> QButtonGroup
        self._widgets_cache = {}         # page_num -> list[Widget], lazy per pagina
        self._field_highlight_items = {}  # page_num -> list[QGraphicsItem], passieve indicator

        # ── Handtekening (fase 5) ──
        self.signature_annotations = []   # list[SignatureAnnotation]
        self._signature_overlay_items = []  # list[_SignatureItem]

        # ── Tekstselectie & hyperlinks (fase 6) ──
        self._selected_text = ""
        self._selection_overlay_items = []  # list[QGraphicsRectItem], blijft staan tot nieuwe sleep/render
        self._links_cache = {}              # page_num -> list[dict], lazy per pagina
        self._hovering_link = False

    def has_unsaved_changes(self):
        return (
            bool(self.text_annotations) or bool(self.highlight_annotations)
            or bool(self.pending_rotations) or bool(self.form_field_values)
            or bool(self.signature_annotations)
        )

    def clear_saved_changes(self):
        """Na succesvol 'Opslaan als': wis de pending-state.

        Tekst-annotaties waren nog geen echte PDF-annotatie, dus hun overlay
        moet ook van het scherm; highlights en rotaties stonden al op het
        live fitz-document en blijven dus gewoon zichtbaar (alleen de
        boekhoudlijsten worden gewist, zoals in NVict_Reader.py:4179-4183).
        Formulierwaarden worden ook gewist (zelfde gedrag als tkinter) - de
        passieve waarde-tekst-indicator verdwijnt daardoor weer, ook al staat
        de waarde inmiddels veilig in het opgeslagen bestand.
        """
        for _annotation, item in self._text_overlay_items:
            self.scene().removeItem(item)
        self._text_overlay_items = []
        for item in self._signature_overlay_items:
            self.scene().removeItem(item)
        self._signature_overlay_items = []
        self.text_annotations = []
        self.highlight_annotations = []
        self.pending_rotations = {}
        self.form_field_values = {}
        self.signature_annotations = []
        self._refresh_field_highlights()

    def rotate_pages(self, page_nums, degrees):
        """Roteer de gegeven pagina's direct op het levende document.

        Poort van NVict_Reader.py:5402-5404 (page.set_rotation, absolute
        hoek). Zichtbaar via de bestaande render, zelfde patroon als
        highlights - en bijgehouden in pending_rotations zodat 'Opslaan als'
        de rotatie kan repliceren (dat heropent het bestand vers vanaf
        schijf, zie save_pdf.py).
        """
        for page_num in page_nums:
            self.pdf_document[page_num].set_rotation(degrees)
            self.pending_rotations[page_num] = degrees
            self._pixmap_cache.pop(page_num, None)
        self.rebuild_layout(force_render=True)

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
        self.pending_rotations = {}
        self.form_mode = False
        self.form_field_values = {}
        self._form_overlay_items = []
        self._radio_groups = {}
        self._widgets_cache = {}
        self._field_highlight_items = {}
        self.signature_annotations = []
        self._signature_overlay_items = []
        self._selected_text = ""
        self._selection_overlay_items = []
        self._links_cache = {}
        self._hovering_link = False

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
        self._form_overlay_items = []
        self._radio_groups = {}
        self._field_highlight_items = {}
        self._signature_overlay_items = []
        # Zelfde reset als tkinter's display_page (tab.selected_text = ""):
        # een volledige her-render (zoom/navigatie) wist de tekstselectie.
        self._selected_text = ""
        self._selection_overlay_items = []

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

        for annotation in self.signature_annotations:
            self._add_signature_overlay(annotation)

        if self.form_mode:
            self._build_form_overlays()

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
                self._remove_field_highlights(page_num)

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
            self._add_field_highlights(page_num)

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
        """Zet de actieve tool ("text_annotate"/"highlight"/None), exclusief.

        Exclusief met form_mode (fase 4): een klik-tool en de aanhoudende
        formulier-overlay-weergave kunnen niet gelijktijdig actief zijn,
        zelfde exclusiviteit als in NVict_Reader.py (punt 7 van het
        fase-4-onderzoek).
        """
        self.tool_mode = mode
        if mode is not None and self.form_mode:
            self.set_form_mode(False)
        if mode == "text_annotate":
            self.setCursor(Qt.CursorShape.IBeamCursor)
        elif mode in ("highlight", "signature"):
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
        if highlight is not None:
            menu = QMenu(self)
            remove_action = menu.addAction("Markering verwijderen")
            chosen = menu.exec(self.viewport().mapToGlobal(self.mapFromScene(scene_pos)))
            if chosen == remove_action:
                self._remove_highlight(highlight)
            return

        signature_hit = self._signature_at(scene_pos)
        if signature_hit is not None:
            annotation, item = signature_hit
            menu = QMenu(self)
            remove_action = menu.addAction("Handtekening verwijderen")
            chosen = menu.exec(self.viewport().mapToGlobal(self.mapFromScene(scene_pos)))
            if chosen == remove_action:
                self._remove_signature(annotation, item)

    # ── Hyperlinks (fase 6) ──────────────────────────────────────────────

    def _links_for_page(self, page_num):
        links = self._links_cache.get(page_num)
        if links is None:
            links = list(self.pdf_document[page_num].get_links())
            self._links_cache[page_num] = links
            rendering.trim_page_cache(self._links_cache, keep=[page_num], anchor=self.current_page,
                                       max_pages=LINKS_CACHE_PAGES)
        return links

    def _link_at_scene_pos(self, scene_pos):
        entry = self._page_at_scene_pos(scene_pos)
        if entry is None:
            return None
        pdf_x = (scene_pos.x() - entry.x) / self.zoom_level
        pdf_y = (scene_pos.y() - entry.y) / self.zoom_level
        for link in self._links_for_page(entry.page_num):
            rect = link.get("from")
            if rect is not None and rect.x0 <= pdf_x <= rect.x1 and rect.y0 <= pdf_y <= rect.y1:
                return link
        return None

    def _handle_link_click(self, link):
        uri = link.get("uri")
        if uri:
            self._open_external_link(uri)
            return
        target_page = link.get("page")
        if target_page is not None and target_page >= 0:
            self.go_to_page(target_page)

    def _open_external_link(self, uri):
        if not security.is_safe_link_url(uri):
            QMessageBox.warning(
                self, "Link geblokkeerd",
                "Deze link is niet geopend omdat het geen gewone web- of e-mailkoppeling is.\n\n"
                "Alleen http, https en mailto worden toegestaan.",
            )
            return

        box = QMessageBox(self)
        box.setWindowTitle("Link openen")
        box.setText("Weet je zeker dat je deze link wilt openen?")
        box.setInformativeText(security.shorten_for_display(uri))
        box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        box.setDefaultButton(QMessageBox.StandardButton.No)
        if box.exec() == QMessageBox.StandardButton.Yes:
            try:
                webbrowser.open(uri)
            except Exception as exc:
                QMessageBox.critical(self, "Fout", f"Kan de link niet openen:\n{exc}")

    # ── Tekstselectie & kopiëren (fase 6) ────────────────────────────────

    def copy_selected_text(self):
        if self._selected_text:
            QApplication.clipboard().setText(self._selected_text)

    def _clear_text_selection(self):
        for item in self._selection_overlay_items:
            self.scene().removeItem(item)
        self._selection_overlay_items = []
        self._selected_text = ""

    def _handle_text_select_release(self, scene_pos):
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

        selected_words = []  # (text, wx0, wy0, wx1, wy1) in scene-coördinaten
        for entry in self.page_layout:
            page_right = entry.x + entry.width
            page_bottom = entry.y + entry.height
            if entry.x > right or page_right < left or entry.y > bottom or page_bottom < top:
                continue
            for word in self._words_for_page(entry.page_num):
                wx0 = entry.x + word.x0 * self.zoom_level
                wy0 = entry.y + word.y0 * self.zoom_level
                wx1 = entry.x + word.x1 * self.zoom_level
                wy1 = entry.y + word.y1 * self.zoom_level
                if wx1 < left or wx0 > right or wy1 < top or wy0 > bottom:
                    continue
                selected_words.append((word.text, wx0, wy0, wx1, wy1))

        if not selected_words:
            return

        # Sorteer leesvolgorde: eerst op regel (y), dan van links naar rechts.
        selected_words.sort(key=lambda w: (round(w[2]), w[1]))
        text_parts = []
        last_y = None
        for text, wx0, wy0, wx1, wy1 in selected_words:
            if last_y is not None and abs(wy0 - last_y) > 5:
                text_parts.append("\n")
            elif text_parts:
                text_parts.append(" ")
            text_parts.append(text)
            last_y = wy0

            rect_item = QGraphicsRectItem(0, 0, wx1 - wx0, wy1 - wy0)
            rect_item.setPos(wx0, wy0)
            rect_item.setBrush(DRAG_SELECTION_BRUSH)
            rect_item.setPen(DRAG_SELECTION_PEN)
            rect_item.setZValue(3)
            self.scene().addItem(rect_item)
            self._selection_overlay_items.append(rect_item)

        self._selected_text = "".join(text_parts).strip()

    # ── Formuliervelden (fase 4) ────────────────────────────────────────

    def _widgets_for_page(self, page_num):
        widgets = self._widgets_cache.get(page_num)
        if widgets is None:
            widgets = list(self.pdf_document[page_num].widgets())
            self._widgets_cache[page_num] = widgets
            rendering.trim_page_cache(self._widgets_cache, keep=[page_num], anchor=self.current_page,
                                       max_pages=FORM_WIDGETS_CACHE_PAGES)
        return widgets

    def _document_has_widgets(self):
        if not self.pdf_document:
            return False
        return any(self._widgets_for_page(page_num) for page_num in range(len(self.pdf_document)))

    def set_form_mode(self, enabled):
        """Schakel de formulier-invulmodus in/uit.

        Geeft False terug als inschakelen niet kan omdat het document geen
        enkel formulierveld heeft (aanroeper toont dan de melding, zoals
        NVict_Reader.py:4223-4226) - form_mode blijft in dat geval uit.
        """
        if enabled:
            if not self._document_has_widgets():
                return False
            self.form_mode = True
            self.set_tool_mode(None)
            self._build_form_overlays()
        else:
            self.form_mode = False
            self._clear_form_overlays()
            self._refresh_field_highlights()
        return True

    def _build_form_overlays(self):
        self._clear_form_overlays()
        for page_num, entry in self._layout_by_page.items():
            for widget in self._widgets_for_page(page_num):
                rect = widget.rect
                x = entry.x + rect.x0 * self.zoom_level
                y = entry.y + rect.y0 * self.zoom_level
                w = max((rect.x1 - rect.x0) * self.zoom_level, 30)
                h = max((rect.y1 - rect.y0) * self.zoom_level, 20)

                field_widget = form_overlay.create_field_widget(widget, self._on_form_value_changed, self._radio_groups)
                field_widget.resize(int(w), int(h))
                proxy = self.scene().addWidget(field_widget)
                proxy.setPos(x, y)
                proxy.setZValue(4)
                self._form_overlay_items.append(proxy)

    def _clear_form_overlays(self):
        for proxy in self._form_overlay_items:
            self.scene().removeItem(proxy)
        self._form_overlay_items = []
        self._radio_groups = {}

    def _on_form_value_changed(self, xref, value):
        self.form_field_values[xref] = value

    def _add_field_highlights(self, page_num):
        entry = self._layout_by_page.get(page_num)
        widgets = self._widgets_for_page(page_num)
        if entry is None or not widgets:
            return
        items = []
        for widget in widgets:
            rect = widget.rect
            x = entry.x + rect.x0 * self.zoom_level
            y = entry.y + rect.y0 * self.zoom_level
            w = (rect.x1 - rect.x0) * self.zoom_level
            h = (rect.y1 - rect.y0) * self.zoom_level

            highlight = QGraphicsRectItem(0, 0, w, h)
            highlight.setPos(x, y)
            highlight.setBrush(FIELD_HIGHLIGHT_BRUSH)
            highlight.setPen(FIELD_HIGHLIGHT_PEN)
            highlight.setZValue(1.5)
            self.scene().addItem(highlight)
            items.append(highlight)

            if not self.form_mode and widget.xref in self.form_field_values:
                value = self.form_field_values[widget.xref]
                if widget.field_type in (form_overlay.FIELD_TYPE_CHECKBOX, form_overlay.FIELD_TYPE_RADIOBUTTON):
                    text = "✓" if value else ""
                else:
                    text = str(value)
                if text:
                    text_item = QGraphicsSimpleTextItem(text)
                    font = QFont("Arial")
                    font.setPixelSize(max(1, round(min(h, 16) * 0.8)))
                    text_item.setFont(font)
                    text_item.setPos(x + 2, y)
                    text_item.setZValue(1.6)
                    self.scene().addItem(text_item)
                    items.append(text_item)

        self._field_highlight_items[page_num] = items

    def _remove_field_highlights(self, page_num):
        for item in self._field_highlight_items.pop(page_num, []):
            self.scene().removeItem(item)

    def _refresh_field_highlights(self):
        for page_num in list(self._pixmap_items.keys()):
            self._remove_field_highlights(page_num)
            self._add_field_highlights(page_num)

    # ── Handtekening (fase 5) ────────────────────────────────────────────

    def _add_signature_overlay(self, annotation: SignatureAnnotation):
        entry = self._layout_by_page.get(annotation.page_num)
        if entry is None:
            return
        image = QImage.fromData(annotation.image_bytes)
        if image.isNull():
            return
        pixmap = QPixmap.fromImage(image).scaled(
            max(1, round(annotation.width * self.zoom_level)),
            max(1, round(annotation.height * self.zoom_level)),
            Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation,
        )
        item = _SignatureItem(pixmap, annotation, self)
        item.setPos(entry.x + annotation.pdf_x * self.zoom_level, entry.y + annotation.pdf_y * self.zoom_level)
        item.setZValue(2)
        self.scene().addItem(item)
        self._signature_overlay_items.append(item)

    def _signature_at(self, scene_pos):
        for item in self._signature_overlay_items:
            if item.contains(item.mapFromScene(scene_pos)):
                return item.annotation, item
        return None

    def _remove_signature(self, annotation, item):
        if annotation in self.signature_annotations:
            self.signature_annotations.remove(annotation)
        if item in self._signature_overlay_items:
            self._signature_overlay_items.remove(item)
        self.scene().removeItem(item)

    def _handle_signature_click(self, scene_pos):
        entry = self._page_at_scene_pos(scene_pos)
        if entry is None:
            return

        dialog = SignatureDialog(self)
        if not dialog.exec():
            return
        image = dialog.get_image()
        if image is None:
            return

        target_width = dialog.get_target_width_pt()
        aspect = image.height() / image.width() if image.width() else 1.0
        target_height = target_width * aspect

        pdf_x = (scene_pos.x() - entry.x) / self.zoom_level
        pdf_y = (scene_pos.y() - entry.y) / self.zoom_level

        annotation = SignatureAnnotation(
            page_num=entry.page_num, pdf_x=pdf_x, pdf_y=pdf_y,
            width=target_width, height=target_height,
            image_bytes=_qimage_to_png_bytes(image),
        )
        self.signature_annotations.append(annotation)
        self._add_signature_overlay(annotation)

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
            if self.tool_mode == "signature":
                # Klik op een bestaande handtekening: laat Qt's ingebouwde
                # item-drag het overnemen (verslepen) i.p.v. een nieuwe te
                # plaatsen.
                hit_item = self.itemAt(event.position().toPoint())
                if isinstance(hit_item, _SignatureItem):
                    super().mousePressEvent(event)
                    return
                self._handle_signature_click(scene_pos)
                return
            if self.tool_mode is None:
                # Normale modus: eerst linkklik checken (stopt daar, net als
                # tkinter's on_click), anders een nieuwe tekstselectie starten.
                link = self._link_at_scene_pos(scene_pos)
                if link is not None:
                    self._clear_text_selection()
                    self._handle_link_click(link)
                    return
                self._clear_text_selection()
                self._handle_highlight_press(scene_pos)  # zelfde sleep-rechthoek-opzet
                return

        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        scene_pos = self.mapToScene(event.position().toPoint())

        if self._drag_rect_item is not None:
            self._handle_highlight_move(scene_pos)
            return

        if self.tool_mode is None and self.pdf_document:
            hovering = self._link_at_scene_pos(scene_pos) is not None
            if hovering != self._hovering_link:
                self._hovering_link = hovering
                if hovering:
                    self.setCursor(Qt.CursorShape.PointingHandCursor)
                else:
                    self.unsetCursor()

        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._drag_rect_item is not None:
            scene_pos = self.mapToScene(event.position().toPoint())
            if self.tool_mode == "highlight":
                self._handle_highlight_release(scene_pos)
            else:
                self._handle_text_select_release(scene_pos)
            return
        super().mouseReleaseEvent(event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.pdf_document and self.zoom_mode == "fit_width":
            self._recompute_fit_width_zoom()
            self.rebuild_layout()
