# -*- coding: utf-8 -*-
"""PDF als bijlage doorsturen per e-mail via Windows Simple MAPI.

Eerdere aanpak (`os.startfile(pad, "sendto")`) faalde met WinError 1155 -
die shell-verb is niet generiek geregistreerd voor .pdf-bestanden. Simple
MAPI (`mapi32.dll`, `MAPISendMail`) is de daadwerkelijke, sinds Windows 95
stabiele API achter "Verzenden naar > E-mailontvanger" en werkt met elk
Simple-MAPI-mailprogramma (Outlook incluis) - via `ctypes`, dus geen
pywin32 nodig.

Bij onopgeslagen wijzigingen wordt eerst een tijdelijke, bijgewerkte PDF
gebouwd (zelfde route als "Opslaan als", save_pdf.build_modified_pdf) -
nooit stilzwijgend het onbewerkte origineel versturen als dat bouwen
faalt.
"""

import ctypes
import os

from PySide6.QtWidgets import QMessageBox

from . import save_pdf
from .i18n import tr

MAPI_LOGON_UI = 0x00000001
MAPI_DIALOG = 0x00000008
SUCCESS_SUCCESS = 0
MAPI_E_USER_ABORT = 1


class MapiFileDesc(ctypes.Structure):
    # Veldvolgorde moet exact overeenkomen met de native MAPI.h-definitie
    # (ulReserved, flFlags, nPosition, lpszPathName, lpszFileName,
    # lpFileType) - een eerdere versie miste `flFlags` en had de velden in
    # de verkeerde volgorde. Omdat MAPISendMail dit geheugen zelf als de
    # ECHTE struct-layout interpreteert, werden lpszFileName en nPosition
    # daardoor verkeerd om ingelezen: nPosition (0xFFFFFFFF) + padding
    # vormde zo een ongeldige pointer 0x00000000FFFFFFFF die het
    # e-mailprogramma probeerde te dereferentiëren - exact de crash die de
    # gebruiker meldde ("access violation reading 0x00000000FFFFFFFF").
    _fields_ = [
        ("ulReserved", ctypes.c_ulong),
        ("flFlags", ctypes.c_ulong),
        ("nPosition", ctypes.c_ulong),
        ("lpszPathName", ctypes.c_char_p),
        ("lpszFileName", ctypes.c_char_p),
        ("lpFileType", ctypes.c_void_p),
    ]


class MapiMessage(ctypes.Structure):
    _fields_ = [
        ("ulReserved", ctypes.c_ulong),
        ("lpszSubject", ctypes.c_char_p),
        ("lpszNoteText", ctypes.c_char_p),
        ("lpszMessageType", ctypes.c_char_p),
        ("lpszDateReceived", ctypes.c_char_p),
        ("lpszConversationID", ctypes.c_char_p),
        ("flFlags", ctypes.c_ulong),
        ("lpOriginator", ctypes.c_void_p),
        ("nRecipCount", ctypes.c_ulong),
        ("lpRecips", ctypes.c_void_p),
        ("nFileCount", ctypes.c_ulong),
        ("lpFiles", ctypes.POINTER(MapiFileDesc)),
    ]


def build_mapi_message(pdf_path, subject=""):
    """Bouw de ctypes-structuren voor MAPISendMail. Los van de DLL-aanroep
    zelf, zodat dit zonder side-effect (geen echt mailvenster) te testen is."""
    file_desc = MapiFileDesc(
        ulReserved=0,
        flFlags=0,
        nPosition=0xFFFFFFFF,
        lpszPathName=os.path.abspath(pdf_path).encode("mbcs"),
        lpszFileName=os.path.basename(pdf_path).encode("mbcs"),
        lpFileType=None,
    )
    message = MapiMessage(
        ulReserved=0,
        lpszSubject=subject.encode("mbcs"),
        lpszNoteText=None,
        lpszMessageType=None,
        lpszDateReceived=None,
        lpszConversationID=None,
        flFlags=0,
        lpOriginator=None,
        nRecipCount=0,
        lpRecips=None,
        nFileCount=1,
        lpFiles=ctypes.pointer(file_desc),
    )
    return message, file_desc  # file_desc levend houden (anders GC'd de pointer)


def call_mapi_send_mail(message) -> int:
    """De daadwerkelijke DLL-aanroep - apart zodat tests dit kunnen stubben."""
    mapi32 = ctypes.windll.mapi32
    return mapi32.MAPISendMail(0, 0, ctypes.byref(message), MAPI_LOGON_UI | MAPI_DIALOG, 0)


def send_as_attachment(parent, tab):
    view = tab.view
    pdf_path = tab.file_path

    if view.has_unsaved_changes():
        try:
            tmp_path = save_pdf.build_modified_pdf(
                tab.file_path, view.text_annotations, view.highlight_annotations,
                view.pending_rotations, view.form_field_values, view.signature_annotations,
            )
        except Exception as exc:
            QMessageBox.critical(
                parent, tr("Wijzigingen niet verwerkt"),
                tr("De wijzigingen konden niet in de PDF worden verwerkt.\n\n"
                   "Details: {error}\n\nEr is niets verstuurd.", error=exc),
            )
            return
        if tmp_path:
            pdf_path = tmp_path

    message, _file_desc = build_mapi_message(pdf_path, subject=os.path.basename(pdf_path))
    try:
        result = call_mapi_send_mail(message)
    except OSError as exc:
        QMessageBox.critical(
            parent, tr("Verzenden mislukt"),
            tr("Kan het e-mailprogramma niet aanroepen.\n\nDetails: {error}", error=exc),
        )
        return

    if result not in (SUCCESS_SUCCESS, MAPI_E_USER_ABORT):
        QMessageBox.critical(
            parent, tr("Verzenden mislukt"),
            tr("Kan geen e-mailprogramma vinden. Controleer of er een standaard-mailprogramma "
               "is ingesteld in Windows.\n\nMAPI-foutcode: {code}", code=result),
        )
