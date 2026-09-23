# -*- coding: utf-8 -*-
"""Zoekdialoog: één tekstveld, Enter of 'Zoeken' start de zoekactie.

Poort van tkinter's show_search_dialog (NVict_Reader.py:2882-2946), als
kleine QDialog i.p.v. een los Toplevel-venster.
"""

from PySide6.QtWidgets import QDialog, QDialogButtonBox, QLabel, QLineEdit, QVBoxLayout


class SearchDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Zoeken in PDF")
        self.setMinimumWidth(340)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Zoek tekst:"))

        self.search_edit = QLineEdit(self)
        self.search_edit.setFocus()
        layout.addWidget(self.search_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Zoeken")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Annuleren")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_search_term(self) -> str:
        return self.search_edit.text().strip()
