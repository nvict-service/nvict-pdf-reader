# -*- coding: utf-8 -*-
"""Dialoog om een tekst-annotatie toe te voegen of te bewerken.

Qt-dialoogvenster in plaats van tkinter's inline canvas-embedded editor
(regel 4640-4880) - past beter bij de wens "modernere dialogen".
"""

from PySide6.QtWidgets import (
    QButtonGroup,
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

from .annotations import COLOR_CHOICES, DEFAULT_FONT_SIZE, FONT_MAP, MAX_FONT_SIZE, MIN_FONT_SIZE

_FONT_NAME_BY_CODE = {v: k for k, v in FONT_MAP.items()}


class TextAnnotationDialog(QDialog):
    """Retourneert via get_result() de ingevoerde annotatiegegevens."""

    def __init__(self, parent=None, existing: dict | None = None):
        super().__init__(parent)
        self.setWindowTitle("Tekst-annotatie bewerken" if existing else "Tekst-annotatie toevoegen")
        self.setMinimumWidth(360)
        self._delete_requested = False
        self._color_value = (existing or {}).get("color", "black")

        layout = QVBoxLayout(self)

        options_row = QHBoxLayout()
        options_row.addWidget(QLabel("Lettertype:"))
        self.font_combo = QComboBox(self)
        self.font_combo.addItems(list(FONT_MAP.keys()))
        self.font_combo.setCurrentText(_FONT_NAME_BY_CODE.get((existing or {}).get("fontname", "helv"), "Helvetica"))
        options_row.addWidget(self.font_combo)

        options_row.addWidget(QLabel("Grootte:"))
        self.size_spin = QSpinBox(self)
        self.size_spin.setRange(MIN_FONT_SIZE, MAX_FONT_SIZE)
        self.size_spin.setValue((existing or {}).get("font_size", DEFAULT_FONT_SIZE))
        options_row.addWidget(self.size_spin)
        layout.addLayout(options_row)

        color_row = QHBoxLayout()
        self.color_group = QButtonGroup(self)
        for label, color_value, bg, fg in COLOR_CHOICES:
            btn = QPushButton(label, self)
            btn.setCheckable(True)
            btn.setStyleSheet(f"background-color: {bg}; color: {fg}; font-weight: bold;")
            btn.setProperty("color_value", color_value)
            if color_value == self._color_value:
                btn.setChecked(True)
            self.color_group.addButton(btn)
            color_row.addWidget(btn)
        layout.addLayout(color_row)

        self.text_edit = QPlainTextEdit(self)
        self.text_edit.setPlainText((existing or {}).get("text", ""))
        self.text_edit.setMinimumHeight(100)
        layout.addWidget(self.text_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        if existing is not None:
            delete_btn = QPushButton("Verwijderen", self)
            delete_btn.clicked.connect(self._on_delete)
            buttons.addButton(delete_btn, QDialogButtonBox.ButtonRole.DestructiveRole)

        layout.addWidget(buttons)

    def _on_delete(self):
        self._delete_requested = True
        self.accept()

    def delete_requested(self) -> bool:
        return self._delete_requested

    def get_result(self) -> dict:
        checked = self.color_group.checkedButton()
        color_value = checked.property("color_value") if checked else "black"
        return {
            "text": self.text_edit.toPlainText().strip(),
            "font_size": self.size_spin.value(),
            "color": color_value,
            "fontname": FONT_MAP.get(self.font_combo.currentText(), "helv"),
        }
