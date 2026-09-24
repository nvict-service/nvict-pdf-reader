# -*- coding: utf-8 -*-
"""Print-dialoogvenster: eigen QDialog i.p.v. het native QPrintDialog.

Een native Windows-printdialoog (QPrintDialog) laat geen eigen velden toe
zoals "passend maken op pagina"/rotatie/kleurmodus - dat zou een tweede
modaal venster nodig maken. Deze dialoog bevat daarom alle opties in één
scherm, zoals de tkinter-versie (regel 3120-3529), maar met Qt-widgets.
"""

from PySide6.QtPrintSupport import QAbstractPrintDialog, QPrintDialog, QPrinter
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QRadioButton,
    QSpinBox,
    QVBoxLayout,
    QComboBox,
    QPushButton,
)

from . import print_backend
from .print_backend import PrintOptions
from .i18n import tr


class PrintDialog(QDialog):
    def __init__(self, parent, document_title, total_pages, current_page):
        super().__init__(parent)
        self.setWindowTitle(tr("PDF Afdrukken"))
        self.setMinimumWidth(420)
        self.total_pages = total_pages
        self.current_page = current_page  # 0-indexed

        self.printer = QPrinter()
        default_name = print_backend.default_printer_name()
        if default_name:
            self.printer.setPrinterName(default_name)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(document_title))

        # ── Printer ──
        printer_row = QHBoxLayout()
        printer_row.addWidget(QLabel(tr("Printer:")))
        self.printer_combo = QComboBox(self)
        names = print_backend.list_printer_names()
        self.printer_combo.addItems(names)
        if default_name in names:
            self.printer_combo.setCurrentText(default_name)
        printer_row.addWidget(self.printer_combo)
        settings_btn = QPushButton(tr("Printer instellingen..."), self)
        settings_btn.clicked.connect(self._open_page_setup)
        printer_row.addWidget(settings_btn)
        layout.addLayout(printer_row)

        # ── Pagina's ──
        pages_box = QGroupBox(tr("Pagina's"), self)
        pages_layout = QVBoxLayout(pages_box)
        self.page_group = QButtonGroup(self)
        self.radio_all = QRadioButton(tr("Alle pagina's ({count})", count=total_pages), self)
        self.radio_all.setChecked(True)
        self.radio_current = QRadioButton(tr("Huidige pagina ({page})", page=current_page + 1), self)
        self.radio_custom = QRadioButton(tr("Aangepast:"), self)
        for rb in (self.radio_all, self.radio_current, self.radio_custom):
            self.page_group.addButton(rb)
            pages_layout.addWidget(rb)
        self.custom_entry = QLineEdit(self)
        self.custom_entry.setPlaceholderText(tr("bv. 1,3,5 of 1-3,5,7-9"))
        pages_layout.addWidget(self.custom_entry)
        layout.addWidget(pages_box)

        # ── Opties ──
        options_box = QGroupBox(tr("Opties"), self)
        options_layout = QVBoxLayout(options_box)
        form = QFormLayout()
        self.copies_spin = QSpinBox(self)
        self.copies_spin.setRange(1, 99)
        self.copies_spin.setValue(1)
        form.addRow(tr("Aantal kopieën:"), self.copies_spin)
        options_layout.addLayout(form)

        self.fit_to_page_check = QCheckBox(tr("Passend maken op pagina"), self)
        self.fit_to_page_check.setChecked(True)
        options_layout.addWidget(self.fit_to_page_check)

        self.duplex_check = QCheckBox(tr("Dubbelzijdig printen"), self)
        options_layout.addWidget(self.duplex_check)
        layout.addWidget(options_box)

        # ── Kleur ──
        color_box = QGroupBox(tr("Kleur"), self)
        color_row = QHBoxLayout(color_box)
        self.color_group = QButtonGroup(self)
        self.radio_color = QRadioButton(tr("Kleur"), self)
        self.radio_color.setChecked(True)
        self.radio_bw = QRadioButton(tr("Zwart-wit"), self)
        for rb in (self.radio_color, self.radio_bw):
            self.color_group.addButton(rb)
            color_row.addWidget(rb)
        layout.addWidget(color_box)

        # ── Rotatie ──
        rotation_box = QGroupBox(tr("Rotatie"), self)
        rotation_row = QHBoxLayout(rotation_box)
        self.rotation_group = QButtonGroup(self)
        self.rotation_values = {}
        for label, value in [("Geen", 0), ("90° rechts", 90), ("180°", 180), ("90° links", 270)]:
            rb = QRadioButton(tr(label), self)
            rb.setChecked(value == 0)
            self.rotation_group.addButton(rb)
            self.rotation_values[rb] = value
            rotation_row.addWidget(rb)
        layout.addWidget(rotation_box)

        # ── Oriëntatie ──
        orientation_box = QGroupBox(tr("Oriëntatie"), self)
        orientation_row = QHBoxLayout(orientation_box)
        self.orientation_group = QButtonGroup(self)
        self.radio_portrait = QRadioButton(tr("Staand"), self)
        self.radio_portrait.setChecked(True)
        self.radio_landscape = QRadioButton(tr("Liggend"), self)
        for rb in (self.radio_portrait, self.radio_landscape):
            self.orientation_group.addButton(rb)
            orientation_row.addWidget(rb)
        layout.addWidget(orientation_box)

        buttons = QDialogButtonBox(self)
        self.print_btn = buttons.addButton(tr("Afdrukken"), QDialogButtonBox.ButtonRole.AcceptRole)
        buttons.addButton(tr("Annuleren"), QDialogButtonBox.ButtonRole.RejectRole)
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _open_page_setup(self):
        """Toon de échte, printer-specifieke instellingen (papierbron,
        afdrukkwaliteit, hardware-duplex, ...).

        QPageSetupDialog (eerder gebruikt) is Qt's generieke, printer-
        onafhankelijke papier/marges-dialoog - Qt schakelt de native
        "Eigenschappen"-knop daar bewust in uit, wat exact de melding was
        die de gebruiker zag. QPrintDialog is wél Qt's eigen mechanisme om
        de driver-eigen instellingen te tonen én correct terug te
        schrijven naar dezelfde QPrinter die we al gebruiken om te
        printen - de pagina-bereik/kopieën-velden daarin schakelen we uit
        omdat onze eigen dialoog daar al in voorziet.
        """
        dialog = QPrintDialog(self.printer, self)
        for option in (
            QAbstractPrintDialog.PrintDialogOption.PrintPageRange,
            QAbstractPrintDialog.PrintDialogOption.PrintCurrentPage,
            QAbstractPrintDialog.PrintDialogOption.PrintSelection,
            QAbstractPrintDialog.PrintDialogOption.PrintCollateCopies,
        ):
            dialog.setOption(option, False)
        dialog.exec()

    def _on_accept(self):
        pages = self._resolve_pages()
        if pages is None:
            return

        if self.duplex_check.isChecked():
            reply = QMessageBox.question(
                self, tr("Dubbelzijdig printen"),
                tr("Dubbelzijdig printen is geselecteerd.\n\n"
                   "Let op: niet alle printers ondersteunen automatisch dubbelzijdig printen.\n\n"
                   "Als uw printer dit niet ondersteunt, ziet u een dialoog om het papier "
                   "handmatig om te draaien.\n\nWilt u doorgaan?"),
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

        self._resolved_pages = pages
        self.accept()

    def _resolve_pages(self):
        if self.radio_current.isChecked():
            return [self.current_page]
        if self.radio_custom.isChecked():
            pages = print_backend.parse_page_range(self.custom_entry.text(), self.total_pages)
            if not pages:
                QMessageBox.critical(
                    self, tr("Ongeldige pagina's"),
                    tr("Ongeldige pagina selectie.\n\nGebruik formaat zoals:\n"
                       "• 1,3,5 (specifieke pagina's)\n• 1-5 (bereik)\n• 1-3,5,7-9 (combinatie)"),
                )
                return None
            return pages
        return list(range(self.total_pages))

    def get_options(self) -> PrintOptions:
        rotation = 0
        for rb, value in self.rotation_values.items():
            if rb.isChecked():
                rotation = value
                break
        return PrintOptions(
            printer_name=self.printer_combo.currentText(),
            pages=self._resolved_pages,
            copies=self.copies_spin.value(),
            fit_to_page=self.fit_to_page_check.isChecked(),
            duplex=self.duplex_check.isChecked(),
            color_mode="kleur" if self.radio_color.isChecked() else "zwart_wit",
            rotation=rotation,
            orientation="staand" if self.radio_portrait.isChecked() else "liggend",
        )
