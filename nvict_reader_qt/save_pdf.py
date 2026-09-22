# -*- coding: utf-8 -*-
"""Wegschrijven van tekst-annotaties en highlights naar een nieuw PDF-bestand.

Poort van _build_modified_pdf/save_changes_to_pdf uit NVict_Reader.py
(regel 3944-4197). Altijd "Opslaan als" - schrijft nooit het origineel
bestand automatisch over, zelfde gedrag als de tkinter-versie.
"""

import os
import shutil
import tempfile

from PySide6.QtWidgets import QFileDialog, QMessageBox

from .annotations import COLOR_MAP
from .document import get_fitz


def build_modified_pdf(file_path, text_annotations, highlight_annotations, pending_rotations=None):
    """Open het originele bestand vers vanaf schijf en voeg wijzigingen toe.

    Geeft het pad naar een tempfile terug, of None als er niets te doen was.
    """
    pending_rotations = pending_rotations or {}
    if not text_annotations and not highlight_annotations and not pending_rotations:
        return None

    fitz = get_fitz()
    doc = fitz.open(file_path)
    try:
        for page_num, degrees in pending_rotations.items():
            doc[page_num].set_rotation(degrees)

        for annotation in text_annotations:
            page = doc[annotation.page_num]
            lines = annotation.text.split("\n")
            max_line_len = max((len(line) for line in lines), default=10)
            approx_width = max_line_len * annotation.font_size * 0.55
            approx_height = len(lines) * annotation.font_size * 1.4
            rect = fitz.Rect(
                annotation.pdf_x, annotation.pdf_y,
                annotation.pdf_x + approx_width + 10,
                annotation.pdf_y + approx_height + 5,
            )
            # BELANGRIJK (regressie uit commit d3a97d7, PyMuPDF >= 1.27):
            # Roep NOOIT annot.set_colors(...) aan NA add_freetext_annot() -
            # dat gooit "cannot be used for FreeText annotations". Kleur en
            # transparantie MOETEN via de constructor-parameters hieronder
            # ingesteld worden.
            annot_obj = page.add_freetext_annot(
                rect, annotation.text,
                fontsize=annotation.font_size, fontname=annotation.fontname,
                text_color=COLOR_MAP.get(annotation.color, (0, 0, 0)),
                fill_color=None, border_color=None, border_width=0,
            )
            annot_obj.update()

        for highlight in highlight_annotations:
            page = doc[highlight.page_num]
            if highlight.quads:
                annot_obj = page.add_highlight_annot(quads=highlight.quads)
                annot_obj.update()

        base_name = os.path.splitext(os.path.basename(file_path))[0]
        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", prefix=f"{base_name}_bewerkt_", delete=False)
        tmp_path = tmp.name
        tmp.close()
        doc.save(tmp_path)
        return tmp_path
    finally:
        doc.close()


def save_as(parent, tab) -> bool:
    """Toon 'Opslaan als', schrijf de annotaties weg. Geeft True bij succes."""
    view = tab.view
    if not view.has_unsaved_changes():
        QMessageBox.information(parent, "Niets te bewaren", "Er zijn geen annotaties om op te slaan.")
        return False

    suggested = os.path.splitext(tab.file_path)[0] + "_bewerkt.pdf"
    target_path, _ = QFileDialog.getSaveFileName(parent, "PDF opslaan als", suggested, "PDF-bestanden (*.pdf)")
    if not target_path:
        return False

    try:
        tmp_path = build_modified_pdf(
            tab.file_path, view.text_annotations, view.highlight_annotations, view.pending_rotations
        )
        if tmp_path is None:
            return False
        shutil.move(tmp_path, target_path)
    except Exception as exc:
        QMessageBox.critical(parent, "Opslaan mislukt", f"Kon het bestand niet opslaan:\n\n{exc}")
        return False

    view.clear_saved_changes()
    QMessageBox.information(parent, "Opgeslagen", f"Opgeslagen als:\n{target_path}")
    return True


def confirm_discard_unsaved(parent, view, action_description) -> bool:
    """Vraag bevestiging als er nog niet-opgeslagen wijzigingen zijn.

    Geeft True terug als er niets te verliezen is, of als de gebruiker
    bevestigt dat hij wil doorgaan. Gebruikt door Exporteren/Samenvoegen
    (regel-3 gat uit het fase-3-onderzoek: die negeerden dit stilzwijgend)
    en door tab/venster sluiten.
    """
    if not view.has_unsaved_changes():
        return True
    reply = QMessageBox.question(
        parent, "Niet-opgeslagen wijzigingen",
        "Er zijn nog niet-opgeslagen tekst-annotaties, markeringen of paginarotaties "
        f"op dit tabblad. Deze worden niet meegenomen in {action_description}.\n\nDoorgaan?",
    )
    return reply == QMessageBox.StandardButton.Yes
