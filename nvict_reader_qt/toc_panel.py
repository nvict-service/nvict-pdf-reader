# -*- coding: utf-8 -*-
"""Inhoudsopgave-paneel: de bladwijzers (outline) die in de PDF zelf staan.

Handleidingen en rapporten hebben er vaak een; een klik springt naar het
hoofdstuk. Zonder inhoudsopgave in de PDF blijft het paneel verborgen.
"""

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QTreeWidget, QTreeWidgetItem

from .i18n import tr

PAGE_ROLE = Qt.ItemDataRole.UserRole


class TocPanel(QTreeWidget):
    def __init__(self, view, parent=None):
        super().__init__(parent)
        self.view = view
        self.setHeaderLabel(tr("Inhoud"))
        self.setMinimumWidth(180)
        self.setUniformRowHeights(True)
        self.itemClicked.connect(self._on_item_clicked)
        self.itemActivated.connect(self._on_item_clicked)

    def rebuild(self):
        self.clear()
        doc = self.view.pdf_document
        if doc is None:
            return
        try:
            toc = doc.get_toc(simple=True)
        except Exception:
            toc = []
        parents = {0: self.invisibleRootItem()}
        for level, title, page in toc:
            parent = parents.get(level - 1) or self.invisibleRootItem()
            item = QTreeWidgetItem(parent, [title.strip() or "…"])
            item.setData(0, PAGE_ROLE, page - 1)  # get_toc telt vanaf 1; -1 = geen doel
            item.setToolTip(0, tr("Pagina {page}", page=page) if page > 0 else title)
            parents[level] = item
            # Diepere niveaus van een vorige tak mogen niet als ouder blijven hangen.
            for deeper in [key for key in parents if key > level]:
                del parents[deeper]
        self.expandToDepth(0)

    def has_entries(self) -> bool:
        return self.topLevelItemCount() > 0

    def _on_item_clicked(self, item, _column=0):
        page = item.data(0, PAGE_ROLE)
        if page is not None and page >= 0:
            self.view.go_to_page(page)
