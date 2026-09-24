# -*- coding: utf-8 -*-
"""Dialoog om tekst op de pagina toe te voegen of te bewerken.

Qt-dialoogvenster in plaats van tkinter's inline canvas-embedded editor
(regel 4640-4880) - past beter bij de wens "modernere dialogen". Kleur
kiezen via QColorDialog i.p.v. 4 vaste knoppen (testfeedback fase 9).
"""

from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QColorDialog,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
)

from .annotations import DEFAULT_FONT_SIZE, DEFAULT_TEXT_COLOR, FONT_MAP, MAX_FONT_SIZE, MIN_FONT_SIZE
from .i18n import tr

_FONT_NAME_BY_CODE = {v: k for k, v in FONT_MAP.items()}


class TextAnnotationDialog(QDialog):
    """Retourneert via get_result() de ingevoerde tekst-gegevens."""

    def __init__(self, parent=None, existing: dict | None = None):
        super().__init__(parent)
        self.setWindowTitle(tr("Tekst bewerken") if existing else tr("Tekst toevoegen"))
        self.setMinimumWidth(360)
        self._delete_requested = False
        self._color = (existing or {}).get("color", DEFAULT_TEXT_COLOR)

        layout = QVBoxLayout(self)

        options_row = QHBoxLayout()
        options_row.addWidget(QLabel(tr("Lettertype:")))
        self.font_combo = QComboBox(self)
        self.font_combo.addItems(list(FONT_MAP.keys()))
        self.font_combo.setCurrentText(_FONT_NAME_BY_CODE.get((existing or {}).get("fontname", "helv"), "Helvetica"))
        options_row.addWidget(self.font_combo)

        options_row.addWidget(QLabel(tr("Grootte:")))
        self.size_spin = QSpinBox(self)
        self.size_spin.setRange(MIN_FONT_SIZE, MAX_FONT_SIZE)
        self.size_spin.setValue((existing or {}).get("font_size", DEFAULT_FONT_SIZE))
        options_row.addWidget(self.size_spin)
        layout.addLayout(options_row)

        color_row = QHBoxLayout()
        color_row.addWidget(QLabel(tr("Kleur:")))
        self.color_btn = QPushButton(tr("Kleur kiezen..."), self)
        self.color_btn.clicked.connect(self._choose_color)
        color_row.addWidget(self.color_btn)
        layout.addLayout(color_row)
        self._update_color_swatch()

        self.text_edit = QPlainTextEdit(self)
        self.text_edit.setPlainText((existing or {}).get("text", ""))
        self.text_edit.setMinimumHeight(100)
        layout.addWidget(self.text_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        if existing is not None:
            delete_btn = QPushButton(tr("Verwijderen"), self)
            delete_btn.clicked.connect(self._on_delete)
            buttons.addButton(delete_btn, QDialogButtonBox.ButtonRole.DestructiveRole)

        layout.addWidget(buttons)

    def _choose_color(self):
        color = QColorDialog.getColor(QColor(self._color), self, tr("Kleur kiezen"))
        if color.isValid():
            self._color = color.name()
            self._update_color_swatch()

    def _update_color_swatch(self):
        fg = "#ffffff" if QColor(self._color).lightness() < 140 else "#000000"
        self.color_btn.setStyleSheet(f"background-color: {self._color}; color: {fg};")

    def _on_delete(self):
        self._delete_requested = True
        self.accept()

    def delete_requested(self) -> bool:
        return self._delete_requested

    def get_result(self) -> dict:
        return {
            "text": self.text_edit.toPlainText().strip(),
            "font_size": self.size_spin.value(),
            "color": self._color,
            "fontname": FONT_MAP.get(self.font_combo.currentText(), "helv"),
        }
