# -*- coding: utf-8 -*-
"""Hoofdvenster: menu, toolbar en tabbladen voor meerdere open PDF's.

Bevat zelf geen documentstate - delegeert naar de actieve tab
(PdfTabWidget), analoog aan hoe NVictReader in de tkinter-versie delegeert
naar self.get_active_tab().
"""

import os

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QIcon, QKeySequence, QPageLayout
from PySide6.QtPrintSupport import QPrinter
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QMainWindow,
    QMessageBox,
    QProgressDialog,
    QTabWidget,
    QToolButton,
)

from . import document_tools, print_backend, save_pdf, settings, theme
from .document_tools_dialogs import ExportPagesDialog, MergePdfsDialog, RotatePagesDialog
from .pdf_tab import PdfTabWidget
from .print_dialog import PrintDialog
from .resources import get_resource_path
from .settings_dialog import SettingsDialog


def _icon(name: str) -> QIcon:
    path = get_resource_path(os.path.join("icons", name))
    return QIcon(path) if os.path.exists(path) else QIcon()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("NVict Reader")
        self.resize(1366, 768)
        self.setMinimumSize(800, 600)
        favicon_path = get_resource_path("favicon.ico")
        if os.path.exists(favicon_path):
            self.setWindowIcon(QIcon(favicon_path))

        self.tabs = QTabWidget(self)
        self.tabs.setTabsClosable(True)
        self.tabs.setMovable(True)
        self.tabs.tabCloseRequested.connect(self._close_tab)
        self.tabs.currentChanged.connect(self._on_tab_changed)
        self.setCentralWidget(self.tabs)

        self._build_actions()
        self._build_menu()
        self._build_toolbar()
        self._update_actions_enabled()

        settings.restore_window_state(self)

    # ── Actions ──────────────────────────────────────────────────────

    def _build_actions(self):
        self.action_open = QAction(_icon("open.png"), "&Openen...", self)
        self.action_open.setShortcut(QKeySequence.StandardKey.Open)
        self.action_open.triggered.connect(self.open_file_dialog)

        self.action_close_tab = QAction("Tab &sluiten", self)
        self.action_close_tab.setShortcut("Ctrl+W")
        self.action_close_tab.triggered.connect(lambda: self._close_tab(self.tabs.currentIndex()))

        self.action_quit = QAction("&Afsluiten", self)
        self.action_quit.setShortcut(QKeySequence.StandardKey.Quit)
        self.action_quit.triggered.connect(self.close)

        self.action_zoom_in = QAction(_icon("zoom-in.png"), "Zoom &in", self)
        self.action_zoom_in.setShortcut("Ctrl+=")
        self.action_zoom_in.triggered.connect(lambda: self._on_active(lambda v: v.zoom_in()))

        self.action_zoom_out = QAction(_icon("zoom-out.png"), "Zoom &uit", self)
        self.action_zoom_out.setShortcut("Ctrl+-")
        self.action_zoom_out.triggered.connect(lambda: self._on_active(lambda v: v.zoom_out()))

        self.action_fit_width = QAction(_icon("fit-width.png"), "&Pasbreedte", self)
        self.action_fit_width.setShortcut("Ctrl+0")
        self.action_fit_width.triggered.connect(lambda: self._on_active(lambda v: v.set_zoom_mode_fit_width()))

        self.action_first_page = QAction(_icon("first-page.png"), "&Eerste pagina", self)
        self.action_first_page.setShortcut("Ctrl+Home")
        self.action_first_page.triggered.connect(lambda: self._on_active(lambda v: v.first_page()))

        self.action_prev_page = QAction(_icon("prev-page.png"), "&Vorige pagina", self)
        self.action_prev_page.setShortcut(QKeySequence.StandardKey.MoveToPreviousPage)
        self.action_prev_page.triggered.connect(lambda: self._on_active(lambda v: v.prev_page()))

        self.action_next_page = QAction(_icon("next-page.png"), "&Volgende pagina", self)
        self.action_next_page.setShortcut(QKeySequence.StandardKey.MoveToNextPage)
        self.action_next_page.triggered.connect(lambda: self._on_active(lambda v: v.next_page()))

        self.action_last_page = QAction(_icon("last-page.png"), "&Laatste pagina", self)
        self.action_last_page.setShortcut("Ctrl+End")
        self.action_last_page.triggered.connect(lambda: self._on_active(lambda v: v.last_page()))

        self.action_print = QAction(_icon("print.png"), "&Afdrukken...", self)
        self.action_print.setShortcut(QKeySequence.StandardKey.Print)
        self.action_print.triggered.connect(self._print_current)

        self.action_save_as = QAction(_icon("save.png"), "&Opslaan als...", self)
        self.action_save_as.setShortcut(QKeySequence.StandardKey.Save)
        self.action_save_as.triggered.connect(self._save_as_current)

        self.action_text_annotate = QAction(_icon("type-text.png"), "&Tekst toevoegen", self)
        self.action_text_annotate.setCheckable(True)
        self.action_text_annotate.toggled.connect(self._on_text_annotate_toggled)

        self.action_highlight = QAction(_icon("marker.png"), "&Markeren", self)
        self.action_highlight.setCheckable(True)
        self.action_highlight.toggled.connect(self._on_highlight_toggled)

        self.action_form_mode = QAction(_icon("form.png"), "&Formulier invullen", self)
        self.action_form_mode.setCheckable(True)
        self.action_form_mode.toggled.connect(self._on_form_mode_toggled)

        self.action_signature = QAction(_icon("check.png"), "&Handtekening plaatsen", self)
        self.action_signature.setCheckable(True)
        self.action_signature.toggled.connect(self._on_signature_toggled)

        self.action_export_pages = QAction(_icon("pages.png"), "Pagina's &exporteren...", self)
        self.action_export_pages.triggered.connect(self._export_pages_current)

        self.action_merge_pdfs = QAction(_icon("copy.png"), "PDF's &samenvoegen...", self)
        self.action_merge_pdfs.triggered.connect(self._merge_pdfs_current)

        self.action_rotate_pages = QAction(_icon("reset.png"), "Pagina &roteren...", self)
        self.action_rotate_pages.triggered.connect(self._rotate_pages_current)

        self.action_copy_text = QAction(_icon("copy.png"), "&Kopiëren", self)
        self.action_copy_text.setShortcut(QKeySequence.StandardKey.Copy)
        self.action_copy_text.triggered.connect(lambda: self._on_active(lambda v: v.copy_selected_text()))

        self.action_settings = QAction("&Instellingen...", self)
        self.action_settings.triggered.connect(self._open_settings)

    def _build_menu(self):
        file_menu = self.menuBar().addMenu("&Bestand")
        file_menu.addAction(self.action_open)
        file_menu.addAction(self.action_close_tab)
        file_menu.addSeparator()
        file_menu.addAction(self.action_print)
        file_menu.addAction(self.action_save_as)
        file_menu.addSeparator()
        file_menu.addAction(self.action_copy_text)
        file_menu.addSeparator()
        file_menu.addAction(self.action_quit)

        view_menu = self.menuBar().addMenu("&Beeld")
        view_menu.addAction(self.action_zoom_in)
        view_menu.addAction(self.action_zoom_out)
        view_menu.addAction(self.action_fit_width)
        view_menu.addSeparator()
        view_menu.addAction(self.action_first_page)
        view_menu.addAction(self.action_prev_page)
        view_menu.addAction(self.action_next_page)
        view_menu.addAction(self.action_last_page)

        annotate_menu = self.menuBar().addMenu("&Annotaties")
        annotate_menu.addAction(self.action_text_annotate)
        annotate_menu.addAction(self.action_highlight)
        annotate_menu.addAction(self.action_form_mode)
        annotate_menu.addAction(self.action_signature)

        edit_menu = self.menuBar().addMenu("&Bewerken")
        edit_menu.addAction(self.action_export_pages)
        edit_menu.addAction(self.action_merge_pdfs)
        edit_menu.addAction(self.action_rotate_pages)
        self.edit_menu = edit_menu

        settings_menu = self.menuBar().addMenu("&Instellingen")
        settings_menu.addAction(self.action_settings)

    def _build_toolbar(self):
        toolbar = self.addToolBar("Hoofdwerkbalk")
        toolbar.setMovable(False)
        toolbar.addAction(self.action_open)
        toolbar.addAction(self.action_print)
        toolbar.addAction(self.action_save_as)
        toolbar.addSeparator()
        toolbar.addAction(self.action_first_page)
        toolbar.addAction(self.action_prev_page)
        toolbar.addAction(self.action_next_page)
        toolbar.addAction(self.action_last_page)
        toolbar.addSeparator()
        toolbar.addAction(self.action_zoom_out)
        toolbar.addAction(self.action_zoom_in)
        toolbar.addAction(self.action_fit_width)
        toolbar.addSeparator()
        toolbar.addAction(self.action_text_annotate)
        toolbar.addAction(self.action_highlight)
        toolbar.addAction(self.action_form_mode)
        toolbar.addAction(self.action_signature)
        toolbar.addSeparator()

        self.edit_menu_button = QToolButton(toolbar)
        self.edit_menu_button.setIcon(_icon("toolbox.png"))
        self.edit_menu_button.setText("Bewerken")
        self.edit_menu_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonIconOnly)
        self.edit_menu_button.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.edit_menu_button.setMenu(self.edit_menu)
        toolbar.addWidget(self.edit_menu_button)

    def _update_actions_enabled(self):
        has_tab = self.tabs.count() > 0
        for action in (
            self.action_close_tab, self.action_zoom_in, self.action_zoom_out,
            self.action_fit_width, self.action_first_page, self.action_prev_page,
            self.action_next_page, self.action_last_page, self.action_print,
            self.action_save_as, self.action_text_annotate, self.action_highlight,
            self.action_form_mode, self.action_signature, self.action_export_pages,
            self.action_merge_pdfs, self.action_rotate_pages, self.action_copy_text,
        ):
            action.setEnabled(has_tab)
        self.edit_menu_button.setEnabled(has_tab)

    # ── Tab helpers ──────────────────────────────────────────────────

    def _on_active(self, func):
        tab = self.tabs.currentWidget()
        if tab is not None:
            func(tab.view)

    def _on_tab_changed(self, _index):
        self._update_actions_enabled()
        tab = self.tabs.currentWidget()
        mode = tab.view.tool_mode if tab is not None else None
        form_active = tab.view.form_mode if tab is not None else False
        actions = (self.action_text_annotate, self.action_highlight, self.action_form_mode, self.action_signature)
        for action in actions:
            action.blockSignals(True)
        self.action_text_annotate.setChecked(mode == "text_annotate")
        self.action_highlight.setChecked(mode == "highlight")
        self.action_signature.setChecked(mode == "signature")
        self.action_form_mode.setChecked(form_active)
        for action in actions:
            action.blockSignals(False)

    def _uncheck_other_tools(self, keep):
        for action in (self.action_text_annotate, self.action_highlight, self.action_form_mode, self.action_signature):
            if action is not keep:
                action.setChecked(False)

    def _on_text_annotate_toggled(self, checked):
        if checked:
            self._uncheck_other_tools(self.action_text_annotate)
        self._apply_tool_mode()

    def _on_highlight_toggled(self, checked):
        if checked:
            self._uncheck_other_tools(self.action_highlight)
        self._apply_tool_mode()

    def _on_signature_toggled(self, checked):
        if checked:
            self._uncheck_other_tools(self.action_signature)
        self._apply_tool_mode()

    def _apply_tool_mode(self):
        mode = None
        if self.action_text_annotate.isChecked():
            mode = "text_annotate"
        elif self.action_highlight.isChecked():
            mode = "highlight"
        elif self.action_signature.isChecked():
            mode = "signature"
        self._on_active(lambda v: v.set_tool_mode(mode))

    def _on_form_mode_toggled(self, checked):
        if checked:
            self.action_text_annotate.setChecked(False)
            self.action_highlight.setChecked(False)
            self.action_signature.setChecked(False)

        tab = self.tabs.currentWidget()
        if tab is None:
            return

        if checked:
            if not tab.view.set_form_mode(True):
                self.action_form_mode.blockSignals(True)
                self.action_form_mode.setChecked(False)
                self.action_form_mode.blockSignals(False)
                QMessageBox.information(
                    self, "Geen formuliervelden", "Dit document heeft geen invulbare formuliervelden."
                )
        else:
            tab.view.set_form_mode(False)

    def _close_tab(self, index):
        if index < 0:
            return
        tab = self.tabs.widget(index)
        if tab is not None and not save_pdf.confirm_discard_unsaved(self, tab.view, "het sluiten van deze tab"):
            return
        self.tabs.removeTab(index)
        if tab is not None:
            tab.close_document()
            tab.deleteLater()
        self._update_actions_enabled()

    # ── File handling ────────────────────────────────────────────────

    def open_file_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "PDF openen", "", "PDF-bestanden (*.pdf)")
        if file_path:
            self.open_file(file_path)

    def open_file(self, file_path):
        try:
            tab = PdfTabWidget(file_path, parent=self.tabs)
        except Exception as exc:
            QMessageBox.critical(self, "Kan PDF niet openen", f"{file_path}\n\n{exc}")
            return
        index = self.tabs.addTab(tab, tab.title)
        self.tabs.setCurrentIndex(index)
        self._update_actions_enabled()

    # ── Printen & opslaan ────────────────────────────────────────────

    def _print_current(self):
        tab = self.tabs.currentWidget()
        if tab is None or not tab.view.pdf_document:
            return

        dialog = PrintDialog(self, tab.title, tab.page_count, tab.view.current_page)
        if not dialog.exec():
            return

        options = dialog.get_options()
        printer = dialog.printer
        printer.setDuplex(QPrinter.DuplexMode.DuplexLongSide if options.duplex else QPrinter.DuplexMode.DuplexNone)
        printer.setPageOrientation(
            QPageLayout.Orientation.Landscape if options.orientation == "liggend" else QPageLayout.Orientation.Portrait
        )

        total_units = len(options.pages) * max(options.copies, 1)
        progress = QProgressDialog("Bezig met printen...", "Annuleren", 0, total_units, self)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)

        def on_progress(done, total):
            progress.setValue(done)
            progress.setLabelText(f"Pagina {done} van {total}")

        try:
            completed = print_backend.run_print_job(
                printer, tab.view.pdf_document, options, on_progress, progress.wasCanceled
            )
        except Exception as exc:
            progress.close()
            QMessageBox.critical(self, "Printfout", f"Kan niet printen:\n\n{exc}")
            return

        progress.close()
        if not completed:
            QMessageBox.information(self, "Geannuleerd", "Het printen is geannuleerd.")

    def _save_as_current(self):
        tab = self.tabs.currentWidget()
        if tab is not None:
            save_pdf.save_as(self, tab)

    # ── Instellingen ──────────────────────────────────────────────────

    def _open_settings(self):
        dialog = SettingsDialog(self)
        if not dialog.exec():
            return
        mode = dialog.get_theme_mode()
        settings.save_theme_mode(mode)
        theme.apply_theme(QApplication.instance(), mode)

    # ── Bewerken-menu: exporteren, samenvoegen, roteren ──────────────

    def _export_pages_current(self):
        tab = self.tabs.currentWidget()
        if tab is None or not tab.view.pdf_document:
            return
        if not save_pdf.confirm_discard_unsaved(self, tab.view, "het geëxporteerde bestand"):
            return

        dialog = ExportPagesDialog(self, tab.page_count)
        if not dialog.exec():
            return
        pages = dialog.get_pages()

        suggested = os.path.splitext(tab.file_path)[0] + "_export.pdf"
        target_path, _ = QFileDialog.getSaveFileName(self, "Pagina's exporteren", suggested, "PDF-bestanden (*.pdf)")
        if not target_path:
            return

        try:
            document_tools.export_pages(tab.view.pdf_document, pages, target_path)
        except Exception as exc:
            QMessageBox.critical(self, "Fout", f"Kan pagina's niet exporteren:\n{exc}")
            return
        QMessageBox.information(self, "Succes", f"{len(pages)} pagina('s) succesvol geëxporteerd naar:\n{target_path}")

    def _merge_pdfs_current(self):
        open_paths = [self.tabs.widget(i).file_path for i in range(self.tabs.count())]
        dialog = MergePdfsDialog(self, open_paths)
        if not dialog.exec():
            return
        file_paths = dialog.get_file_paths()

        target_path, _ = QFileDialog.getSaveFileName(self, "PDF's samenvoegen", "samengevoegd.pdf", "PDF-bestanden (*.pdf)")
        if not target_path:
            return

        try:
            document_tools.merge_pdfs(file_paths, target_path)
        except Exception as exc:
            QMessageBox.critical(self, "Fout", f"Kan PDF's niet samenvoegen:\n{exc}")
            return

        reply = QMessageBox.question(
            self, "Succes",
            f"{len(file_paths)} bestanden succesvol samengevoegd naar:\n{target_path}\n\nNu openen?",
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.open_file(target_path)

    def _rotate_pages_current(self):
        tab = self.tabs.currentWidget()
        if tab is None or not tab.view.pdf_document:
            return

        dialog = RotatePagesDialog(self, tab.page_count, tab.view.current_page)
        if not dialog.exec():
            return
        pages = dialog.get_pages()
        degrees = dialog.get_rotation()

        try:
            tab.view.rotate_pages(pages, degrees)
        except Exception as exc:
            QMessageBox.critical(self, "Fout", f"Kan pagina's niet roteren:\n{exc}")
            return
        QMessageBox.information(
            self, "Geroteerd",
            f"{len(pages)} pagina('s) geroteerd met {degrees}°.\n\nVergeet niet op te slaan om de wijziging te behouden!",
        )

    # ── Window lifecycle ─────────────────────────────────────────────

    def closeEvent(self, event):
        for index in range(self.tabs.count()):
            widget = self.tabs.widget(index)
            if widget is not None and not save_pdf.confirm_discard_unsaved(self, widget.view, "het afsluiten van NVict Reader"):
                event.ignore()
                return

        settings.save_window_state(self)
        for index in range(self.tabs.count()):
            widget = self.tabs.widget(index)
            if widget is not None:
                widget.close_document()
        super().closeEvent(event)
