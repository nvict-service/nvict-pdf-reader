# -*- coding: utf-8 -*-
"""Dialogen voor het Bewerken-menu: pagina's exporteren, PDF's samenvoegen,
pagina's roteren. Poort van de gelijknamige tkinter-Toplevels (regel
5189-5433, 5645-5829).
"""

import os

from PySide6.QtWidgets import (
    QButtonGroup,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
)

from . import print_backend


class ExportPagesDialog(QDialog):
    def __init__(self, parent, total_pages):
        super().__init__(parent)
        self.setWindowTitle("Pagina's Exporteren")
        self.total_pages = total_pages

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"Document heeft {total_pages} pagina('s)."))
        layout.addWidget(QLabel("Welke pagina's wilt u exporteren?"))

        self.entry = QLineEdit(self)
        self.entry.setText(f"1-{total_pages}")
        self.entry.setPlaceholderText("bv. 1,3,5 of 1-5 of 1-3,5,7-9")
        layout.addWidget(self.entry)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Exporteren")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Annuleren")
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self._resolved_pages = None

    def _on_accept(self):
        pages = print_backend.parse_page_range(self.entry.text(), self.total_pages)
        if not pages:
            QMessageBox.critical(
                self, "Ongeldige pagina's",
                "Ongeldige pagina selectie.\n\nGebruik formaat zoals:\n"
                "• 1,3,5 (specifieke pagina's)\n• 1-5 (bereik)\n• 1-3,5,7-9 (combinatie)",
            )
            return
        self._resolved_pages = pages
        self.accept()

    def get_pages(self):
        return self._resolved_pages


class RotatePagesDialog(QDialog):
    def __init__(self, parent, total_pages, current_page):
        super().__init__(parent)
        self.setWindowTitle("Pagina's Roteren")
        self.total_pages = total_pages

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Welke pagina's?"))
        self.entry = QLineEdit(self)
        self.entry.setText(str(current_page + 1))
        self.entry.setPlaceholderText("bv. 1,3,5 of 1-5")
        layout.addWidget(self.entry)

        layout.addWidget(QLabel("Rotatie:"))
        self.rotation_group = QButtonGroup(self)
        self.rotation_buttons = {}
        for angle in (90, 180, 270):
            label = f"{angle}° (rechtsom)" if angle == 90 else f"{angle}°"
            rb = QRadioButton(label, self)
            rb.setChecked(angle == 90)
            self.rotation_group.addButton(rb)
            self.rotation_buttons[rb] = angle
            layout.addWidget(rb)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Roteren")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Annuleren")
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self._resolved_pages = None

    def _on_accept(self):
        pages = print_backend.parse_page_range(self.entry.text(), self.total_pages)
        if not pages:
            QMessageBox.critical(self, "Ongeldige invoer", "Ongeldige pagina selectie!")
            return
        self._resolved_pages = pages
        self.accept()

    def get_pages(self):
        return self._resolved_pages

    def get_rotation(self):
        for rb, angle in self.rotation_buttons.items():
            if rb.isChecked():
                return angle
        return 90


class MergePdfsDialog(QDialog):
    def __init__(self, parent, open_tab_paths):
        super().__init__(parent)
        self.setWindowTitle("PDF's Samenvoegen")
        self.setMinimumSize(420, 380)
        self.open_tab_paths = open_tab_paths
        self.file_paths = []

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Bestanden om samen te voegen (in deze volgorde):"))

        self.list_widget = QListWidget(self)
        layout.addWidget(self.list_widget)

        button_row = QHBoxLayout()
        add_open_btn = QPushButton("Open tabbladen toevoegen", self)
        add_open_btn.clicked.connect(self._add_open_tabs)
        button_row.addWidget(add_open_btn)
        add_btn = QPushButton("Toevoegen...", self)
        add_btn.clicked.connect(self._add_files)
        button_row.addWidget(add_btn)
        remove_btn = QPushButton("Verwijderen", self)
        remove_btn.clicked.connect(self._remove_selected)
        button_row.addWidget(remove_btn)
        up_btn = QPushButton("Omhoog", self)
        up_btn.clicked.connect(lambda: self._move_selected(-1))
        button_row.addWidget(up_btn)
        down_btn = QPushButton("Omlaag", self)
        down_btn.clicked.connect(lambda: self._move_selected(1))
        button_row.addWidget(down_btn)
        layout.addLayout(button_row)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Combineren")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Annuleren")
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _add_open_tabs(self):
        for path in self.open_tab_paths:
            if path not in self.file_paths:
                self.file_paths.append(path)
        self._refresh_list()

    def _add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "PDF-bestanden toevoegen", "", "PDF-bestanden (*.pdf)")
        for path in paths:
            if path not in self.file_paths:
                self.file_paths.append(path)
        self._refresh_list()

    def _remove_selected(self):
        row = self.list_widget.currentRow()
        if row >= 0:
            self.file_paths.pop(row)
            self._refresh_list()

    def _move_selected(self, delta):
        row = self.list_widget.currentRow()
        new_row = row + delta
        if row < 0 or not (0 <= new_row < len(self.file_paths)):
            return
        self.file_paths[row], self.file_paths[new_row] = self.file_paths[new_row], self.file_paths[row]
        self._refresh_list()
        self.list_widget.setCurrentRow(new_row)

    def _refresh_list(self):
        self.list_widget.clear()
        for path in self.file_paths:
            self.list_widget.addItem(os.path.basename(path))

    def _on_accept(self):
        if len(self.file_paths) < 2:
            QMessageBox.warning(self, "Te weinig bestanden", "Voeg minstens 2 PDF-bestanden toe om samen te voegen.")
            return
        self.accept()

    def get_file_paths(self):
        return list(self.file_paths)
