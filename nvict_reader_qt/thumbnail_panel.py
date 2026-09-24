# -*- coding: utf-8 -*-
"""Paginathumbnails-paneel, toggle-baar via de werkbalk/Beeld-menu.

Poort van tkinter's thumbnail-paneel (NVict_Reader.py regel 945-984 opbouw,
6284-6410 rendering/interactie), maar met lazy rendering per zichtbare rij
in plaats van het hele document vooraf in een achtergrondthread te
renderen - vermijdt een PyMuPDF-thread-veiligheidsrisico (hetzelfde
fitz-document wordt anders gelijktijdig vanuit twee threads benaderd) en
sluit aan bij de bestaande lazy-cache-aanpak in pdf_view.py.
"""

from PySide6.QtCore import QSize, Qt, QTimer
from PySide6.QtGui import QIcon, QImage, QPixmap
from PySide6.QtWidgets import QListView, QListWidget, QListWidgetItem

from .document import get_fitz
from .i18n import tr

THUMBNAIL_WIDTH = 120
RENDER_DEBOUNCE_MS = 60
VISIBLE_MARGIN_ROWS = 2


class ThumbnailPanel(QListWidget):
    def __init__(self, view, parent=None):
        super().__init__(parent)
        self.view = view
        self._rendered_rows = set()

        self.setFixedWidth(THUMBNAIL_WIDTH + 40)
        self.setViewMode(QListView.ViewMode.IconMode)
        self.setFlow(QListView.Flow.TopToBottom)
        self.setWrapping(False)
        self.setResizeMode(QListView.ResizeMode.Adjust)
        self.setMovement(QListView.Movement.Static)
        self.setIconSize(QSize(THUMBNAIL_WIDTH, int(THUMBNAIL_WIDTH * 1.7)))
        self.setSpacing(3)
        self.setUniformItemSizes(False)

        self._render_timer = QTimer(self)
        self._render_timer.setSingleShot(True)
        self._render_timer.timeout.connect(self._render_visible_rows)

        self.itemClicked.connect(self._on_item_clicked)
        self.verticalScrollBar().valueChanged.connect(self._schedule_render)
        view.pageChanged.connect(self._on_page_changed)

    def rebuild(self):
        """Bouw de itemslijst opnieuw op (bij een nieuw/gesloten document)."""
        self.clear()
        self._rendered_rows = set()
        if not self.view.pdf_document:
            return
        for page_num in range(len(self.view.pdf_document)):
            item = QListWidgetItem(tr("Pagina {page}", page=page_num + 1))
            item.setData(Qt.ItemDataRole.UserRole, page_num)
            item.setTextAlignment(Qt.AlignmentFlag.AlignHCenter)
            self.addItem(item)
        self._schedule_render()

    def _schedule_render(self, *_args):
        self._render_timer.start(RENDER_DEBOUNCE_MS)

    def _render_visible_rows(self):
        if not self.view.pdf_document or self.count() == 0:
            return

        # Direct na het openen van een document (paneel nu standaard
        # zichtbaar) kan deze timer afgaan vóórdat Qt de lay-out van het
        # net getoonde paneel heeft bijgewerkt - viewport en/of item-
        # posities zijn dan nog (0,0). Zonder deze check bleef het paneel
        # dan stil helemaal leeg (pas een handmatige aan/uit-toggle, die
        # via showEvent opnieuw rendert wanneer de lay-out wél klaar is,
        # herstelde het) - nu proberen we het gewoon nog een keer.
        if self.viewport().height() <= 0:
            self._schedule_render()
            return

        # indexAt() op de rand-punten van de viewport is onbetrouwbaar zodra
        # een punt net in de marge/spacing tussen items valt (geeft dan -1,
        # wat zonder deze aanpak per ongeluk "render alles" zou triggeren).
        # Een rect-intersectie per item is een goedkope check (geen render)
        # en blijft correct ongeacht spacing/marges.
        first_item_rect = self.visualItemRect(self.item(0))
        if first_item_rect.height() <= 0:
            self._schedule_render()
            return
        row_height = max(1, first_item_rect.height() + self.spacing())
        margin_px = VISIBLE_MARGIN_ROWS * row_height
        viewport_rect = self.viewport().rect().adjusted(0, -margin_px, 0, margin_px)

        any_rendered = False
        for row in range(self.count()):
            if row in self._rendered_rows:
                continue
            item = self.item(row)
            if item is None:
                continue
            if not self.visualItemRect(item).intersects(viewport_rect):
                continue
            item.setIcon(QIcon(self._render_thumbnail(row)))
            self._rendered_rows.add(row)
            any_rendered = True

        if any_rendered:
            # Items werden via addItem() zonder icoon toegevoegd, dus Qt's
            # gecachete sizeHint (in niet-uniforme icoonmodus) was aanvankelijk
            # alleen op de tekst gebaseerd - te laag voor de echte thumbnail.
            # Zonder deze expliciete her-layout bleef elke rij daardoor
            # afgeknipt tot enkel het bovenste randje van de afbeelding
            # (zichtbaar als een herhaald "bannertje"), totdat een
            # handmatige aan/uit-toggle toevallig een volledige her-layout
            # forceerde.
            self.doItemsLayout()

    def _render_thumbnail(self, page_num) -> QPixmap:
        fitz = get_fitz()
        page = self.view.pdf_document[page_num]
        bound = page.bound()
        zoom = THUMBNAIL_WIDTH / bound.width if bound.width else 1.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        fmt = QImage.Format.Format_RGBA8888 if pix.alpha else QImage.Format.Format_RGB888
        image = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
        return QPixmap.fromImage(image.copy())

    def _on_item_clicked(self, item):
        page_num = item.data(Qt.ItemDataRole.UserRole)
        if page_num is not None:
            self.view.go_to_page(page_num)

    def _on_page_changed(self, page_num):
        self.blockSignals(True)
        self.setCurrentRow(page_num)
        self.blockSignals(False)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._schedule_render()

    def showEvent(self, event):
        super().showEvent(event)
        self._schedule_render()
