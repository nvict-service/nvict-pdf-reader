# -*- coding: utf-8 -*-
"""Exporteren van pagina's en samenvoegen van PDF's.

Poort van export_pages/merge_pdfs uit NVict_Reader.py (regel 5189-5328,
5645-5829). De vrijwel identieke, nooit aan een menu gekoppelde legacy-
functie extract_pages (regel 5517+) wordt bewust niet overgenomen.
"""

from .document import get_fitz


def export_pages(pdf_document, page_nums, save_path):
    """Sla de gegeven 0-indexed paginanummers op als een nieuw PDF-bestand."""
    fitz = get_fitz()
    new_doc = fitz.open()
    try:
        for page_num in page_nums:
            new_doc.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)
        new_doc.save(save_path)
    finally:
        new_doc.close()


def merge_pdfs(file_paths, save_path, passwords=None):
    """Voeg meerdere PDF-bestanden samen (in de gegeven volgorde) tot één bestand.

    `passwords`: pad -> wachtwoord voor beveiligde bestanden (zie
    MainWindow._password_for_file, dat ze vooraf bij de gebruiker ophaalt).
    """
    fitz = get_fitz()
    passwords = passwords or {}
    merged_doc = fitz.open()
    try:
        for file_path in file_paths:
            pdf_doc = fitz.open(file_path)
            if pdf_doc.needs_pass or passwords.get(file_path):
                pdf_doc.authenticate(passwords.get(file_path) or "")
            try:
                merged_doc.insert_pdf(pdf_doc)
            finally:
                pdf_doc.close()
        merged_doc.save(save_path)
    finally:
        merged_doc.close()


def protect_pdf(source_path, target_path, open_password="", owner_password="",
                allow_print=True, allow_copy=False, allow_annotate=False, source_password=None):
    """Sla `source_path` beveiligd (AES-256) op als `target_path`.

    `open_password` is nodig om de PDF te openen; `owner_password` beperkt
    wat er met de geopende PDF mag (de allow_*-vlaggen). Zonder
    eigenaarswachtwoord maar mét openingswachtwoord gebruikt de PDF dat
    openingswachtwoord ook als eigenaarswachtwoord: wie hem kan openen, mag
    dan alles. Schrijft eerst naar een tijdelijk bestand naast het doel,
    zodat ook het momenteel geopende bestand zelf als doel kan dienen.
    """
    import os
    import tempfile

    fitz = get_fitz()
    if owner_password:
        permissions = fitz.PDF_PERM_ACCESSIBILITY
        if allow_print:
            permissions |= fitz.PDF_PERM_PRINT | fitz.PDF_PERM_PRINT_HQ
        if allow_copy:
            permissions |= fitz.PDF_PERM_COPY
        if allow_annotate:
            permissions |= fitz.PDF_PERM_ANNOTATE | fitz.PDF_PERM_FORM
    else:
        permissions = -1  # alles toegestaan

    doc = fitz.open(source_path)
    try:
        if doc.needs_pass:
            doc.authenticate(source_password or "")
        handle, tmp_path = tempfile.mkstemp(suffix=".pdf", dir=os.path.dirname(os.path.abspath(target_path)))
        os.close(handle)
        try:
            doc.save(
                tmp_path,
                encryption=fitz.PDF_ENCRYPT_AES_256,
                user_pw=open_password,
                owner_pw=owner_password or open_password,
                permissions=permissions,
                garbage=3,
                deflate=True,
            )
        except Exception:
            os.remove(tmp_path)
            raise
    finally:
        doc.close()
    os.replace(tmp_path, target_path)


def save_without_protection(source_path, target_path, password=None):
    """Sla een kopie op zonder wachtwoord en zonder beperkingen. Alleen
    aanroepen nadat de gebruiker het eigenaarswachtwoord heeft ingevoerd."""
    import os
    import tempfile

    fitz = get_fitz()
    doc = fitz.open(source_path)
    try:
        if doc.needs_pass or password:
            doc.authenticate(password or "")
        handle, tmp_path = tempfile.mkstemp(suffix=".pdf", dir=os.path.dirname(os.path.abspath(target_path)))
        os.close(handle)
        try:
            doc.save(tmp_path, encryption=fitz.PDF_ENCRYPT_NONE, garbage=3, deflate=True)
        except Exception:
            os.remove(tmp_path)
            raise
    finally:
        doc.close()
    os.replace(tmp_path, target_path)


# ── Paginabewerkingen ─────────────────────────────────────────────────
# Een bewerking is een tuple, zodat hij bij "Opslaan als" op het verse
# bestand van schijf exact herhaald kan worden:
#   ("delete", [paginanummers])
#   ("move", van, naar)              - "naar" is de nieuwe positie (0-based)
#   ("insert", pad, wachtwoord, bij) - alle pagina's van pad vóór index "bij"


def move_page_to(doc, source, target):
    """Verplaats pagina `source` zodat hij daarna op index `target` staat."""
    if source == target:
        return
    if target >= len(doc) - 1:
        doc.move_page(source, -1)
    elif target > source:
        doc.move_page(source, target + 1)
    else:
        doc.move_page(source, target)


def apply_page_op(doc, op):
    kind = op[0]
    if kind == "delete":
        doc.delete_pages(sorted(op[1]))
    elif kind == "move":
        move_page_to(doc, op[1], op[2])
    elif kind == "insert":
        _kind, path, password, at = op
        source = get_fitz().open(path)
        try:
            if source.needs_pass or password:
                source.authenticate(password or "")
            doc.insert_pdf(source, start_at=at)
            return len(source)
        finally:
            source.close()
    else:
        raise ValueError(f"Onbekende paginabewerking: {kind}")
    return None


# Niveau -> (doel-dpi voor afbeeldingen, JPEG-kwaliteit)
COMPRESS_LEVELS = {
    "light": (200, 85),
    "medium": (150, 75),
    "strong": (100, 60),
}


def compress_pdf(source_path, target_path, level="medium", password=None):
    """Verklein een PDF: afbeeldingen naar lagere resolutie/JPEG-kwaliteit
    en ongebruikte objecten opruimen. Tekst en vectorafbeeldingen blijven
    scherp. Geeft (oude grootte, nieuwe grootte) in bytes terug; is het
    resultaat niet kleiner, dan wordt er niets geschreven."""
    import os
    import tempfile

    fitz = get_fitz()
    dpi, quality = COMPRESS_LEVELS[level]
    original_size = os.path.getsize(source_path)
    doc = fitz.open(source_path)
    try:
        if doc.needs_pass or password:
            doc.authenticate(password or "")
        doc.rewrite_images(dpi_threshold=dpi + 10, dpi_target=dpi, quality=quality)
        handle, tmp_path = tempfile.mkstemp(suffix=".pdf", dir=os.path.dirname(os.path.abspath(target_path)))
        os.close(handle)
        try:
            doc.save(
                tmp_path, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True,
                clean=True, use_objstms=1, encryption=fitz.PDF_ENCRYPT_NONE,
            )
        except Exception:
            os.remove(tmp_path)
            raise
    finally:
        doc.close()
    new_size = os.path.getsize(tmp_path)
    if new_size >= original_size:
        os.remove(tmp_path)
        return original_size, new_size
    os.replace(tmp_path, target_path)
    return original_size, new_size
