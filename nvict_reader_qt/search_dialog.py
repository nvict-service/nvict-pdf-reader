# -*- coding: utf-8 -*-
"""Zoekdialoog: één tekstveld, Enter of 'Zoeken' start de zoekactie.

Poort van tkinter's show_search_dialog (NVict_Reader.py:2882-2946), als
kleine QDialog i.p.v. een los Toplevel-venster.
"""

from PySide6.QtWidgets import QDialog, QDialogButtonBox, QLabel, QLineEdit, QVBoxLayout

from .i18n import tr


class SearchDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("Zoeken in PDF"))
        self.setMinimumWidth(340)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(tr("Zoek tekst:")))

        self.search_edit = QLineEdit(self)
        self.search_edit.setFocus()
        layout.addWidget(self.search_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText(tr("Zoeken"))
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText(tr("Annuleren"))
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_search_term(self) -> str:
        return self.search_edit.text().strip()
