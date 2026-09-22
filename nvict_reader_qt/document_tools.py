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


def merge_pdfs(file_paths, save_path):
    """Voeg meerdere PDF-bestanden samen (in de gegeven volgorde) tot één bestand."""
    fitz = get_fitz()
    merged_doc = fitz.open()
    try:
        for file_path in file_paths:
            pdf_doc = fitz.open(file_path)
            try:
                merged_doc.insert_pdf(pdf_doc)
            finally:
                pdf_doc.close()
        merged_doc.save(save_path)
    finally:
        merged_doc.close()
