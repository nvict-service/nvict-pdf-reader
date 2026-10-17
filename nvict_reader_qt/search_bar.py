# -*- coding: utf-8 -*-
"""Zoekbalk onderaan het tabblad (Ctrl+F): alle treffers in het document,
met vorige/volgende (Enter / Shift+Enter, F3 / Shift+F3) en "3 van 12".
Vervangt het oude zoekvenster dat alleen de eerste treffer toonde."""

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QHBoxLayout, QLabel, QLineEdit, QToolButton, QWidget

from .i18n import tr


class SearchBar(QWidget):
    closed = Signal()

    def __init__(self, view, parent=None):
        super().__init__(parent)
        self.view = view
        self._term = ""

        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 4, 8, 4)
        layout.addWidget(QLabel(tr("Zoeken:"), self))

        self.edit = QLineEdit(self)
        self.edit.setPlaceholderText(tr("Zoek in document..."))
        self.edit.setClearButtonEnabled(True)
        self.edit.setMaximumWidth(320)
        self.edit.returnPressed.connect(self._on_return)
        layout.addWidget(self.edit)

        self.prev_btn = QToolButton(self)
        self.prev_btn.setText("▲")
        self.prev_btn.setToolTip(tr("Vorige treffer (Shift+Enter)"))
        self.prev_btn.clicked.connect(self.find_previous)
        layout.addWidget(self.prev_btn)

        self.next_btn = QToolButton(self)
        self.next_btn.setText("▼")
        self.next_btn.setToolTip(tr("Volgende treffer (Enter)"))
        self.next_btn.clicked.connect(self.find_next)
        layout.addWidget(self.next_btn)

        self.count_label = QLabel(self)
        self.count_label.setMinimumWidth(120)
        layout.addWidget(self.count_label)
        layout.addStretch()

        close_btn = QToolButton(self)
        close_btn.setText("✕")
        close_btn.setToolTip(tr("Zoekbalk sluiten (Esc)"))
        close_btn.clicked.connect(self.close_bar)
        layout.addWidget(close_btn)

    def open_bar(self):
        self.show()
        selected = self.view._selected_text.strip()
        if selected and "\n" not in selected and len(selected) < 80:
            self.edit.setText(selected)
        self.edit.setFocus()
        self.edit.selectAll()

    def close_bar(self):
        self.view.clear_search()
        self._term = ""
        self.count_label.clear()
        self.hide()
        self.view.setFocus()
        self.closed.emit()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.close_bar()
            return
        super().keyPressEvent(event)

    def _on_return(self):
        from PySide6.QtWidgets import QApplication
        if QApplication.keyboardModifiers() & Qt.KeyboardModifier.ShiftModifier:
            self.find_previous()
        else:
            self.find_next()

    def _ensure_searched(self) -> bool:
        term = self.edit.text().strip()
        if not term:
            self.view.clear_search()
            self._term = ""
            self.count_label.clear()
            return False
        if term != self._term:
            self._term = term
            self.view.search_all(term)
        return True

    def find_next(self):
        if self._ensure_searched():
            self._show(self.view.next_search_hit(+1))

    def find_previous(self):
        if self._ensure_searched():
            self._show(self.view.next_search_hit(-1))

    def _show(self, index):
        total = len(self.view.search_hits)
        if total == 0:
            self.count_label.setText(tr("Niet gevonden"))
        else:
            self.count_label.setText(tr("{current} van {total}", current=index + 1, total=total))
