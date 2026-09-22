# -*- coding: utf-8 -*-
"""Beveiligingshelpers, overgenomen uit NVict_Reader.py (regel 48-88).

Geen tkinter-afhankelijkheid in de brondefinities, dus verbatim overgenomen.
"""

import urllib.parse

# Schemes die we uit een PDF-link mogen doorgeven aan de browser. Alles daar
# buiten (file:, ms-msdt:, javascript:, een UNC-pad, ...) wordt op Windows door
# de shell uitgevoerd en zou een kwaadaardige PDF een startknop geven.
ALLOWED_LINK_SCHEMES = ("http", "https", "mailto")

# Updates mogen uitsluitend van onze eigen site komen. De download-URL staat in
# een JSON-bestand op de server; zonder deze controle zou een aangepaste of
# onderschepte JSON de installer van een willekeurige host kunnen halen.
UPDATE_ALLOWED_HOSTS = ("www.nvict.nl", "nvict.nl")


def is_safe_link_url(url):
    """Bepaal of een URL uit een PDF veilig aan de browser doorgegeven mag worden."""
    if not url or not isinstance(url, str):
        return False
    # Regeleindes/tabs kunnen gebruikt worden om de weergegeven URL te vervalsen
    if any(c in url for c in ("\n", "\r", "\t", "\x00")):
        return False
    try:
        scheme = urllib.parse.urlparse(url).scheme.lower()
    except Exception:
        return False
    if scheme not in ALLOWED_LINK_SCHEMES:
        return False
    # "http:evil.exe" o.i.d. zonder net-locatie afwijzen (mailto heeft er geen)
    if scheme in ("http", "https") and not url.lower().startswith(scheme + "://"):
        return False
    return True


def is_trusted_update_url(url):
    """Bepaal of een download-URL voor een update van onze eigen server komt."""
    if not url or not isinstance(url, str):
        return False
    try:
        parts = urllib.parse.urlparse(url)
    except Exception:
        return False
    # Alleen https: anders kan een man-in-the-middle de installer vervangen
    if parts.scheme.lower() != "https":
        return False
    return parts.hostname is not None and parts.hostname.lower() in UPDATE_ALLOWED_HOSTS


def shorten_for_display(text, limit=180):
    """Kort tekst in voor weergave in een dialoog met een vaste hoogte."""
    text = "".join(ch for ch in str(text) if ch.isprintable())
    if len(text) > limit:
        return text[:limit] + "…"
    return text
