# -*- coding: utf-8 -*-
"""Taalkeuze Nederlands/Engels.

De Nederlandse tekst in de code is de sleutel: tr("Openen") geeft in het
Engels "Open" terug (uit i18n_en.EN) en in het Nederlands gewoon "Openen".
Een tekst zonder Engelse vertaling blijft dus Nederlands in plaats van leeg,
en de code blijft net zo leesbaar als voorheen.

Placeholders gaan via str.format: tr("Pagina {n} van {total}", n=1, total=5).

De taal wordt één keer bij het opstarten gekozen (init). Een wijziging in
Instellingen geldt na een herstart van het programma.
"""

from PySide6.QtCore import QLocale

from .i18n_en import EN

LANGUAGE_SYSTEM = "system"
LANGUAGE_DUTCH = "nl"
LANGUAGE_ENGLISH = "en"

_language = LANGUAGE_DUTCH


def system_language() -> str:
    """Nederlands als Windows in het Nederlands staat, anders Engels."""
    return LANGUAGE_DUTCH if QLocale.system().language() == QLocale.Language.Dutch else LANGUAGE_ENGLISH


def resolve(setting: str) -> str:
    if setting in (LANGUAGE_DUTCH, LANGUAGE_ENGLISH):
        return setting
    return system_language()


def init(setting: str) -> str:
    global _language
    _language = resolve(setting)
    return _language


def current() -> str:
    return _language


def tr(text: str, **kwargs) -> str:
    if _language == LANGUAGE_ENGLISH:
        text = EN.get(text, text)
    return text.format(**kwargs) if kwargs else text
