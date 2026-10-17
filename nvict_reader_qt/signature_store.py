# -*- coding: utf-8 -*-
"""Bewaarde handtekeningen, zodat de gebruiker ze niet elke keer opnieuw
hoeft te tekenen. Als PNG in de lokale app-gegevensmap van deze gebruiker
(niet in het register: dat is voor afbeeldingen niet bedoeld)."""

import hashlib
import os

from PySide6.QtCore import QByteArray, QBuffer, QIODevice, QStandardPaths
from PySide6.QtGui import QImage

MAX_SAVED = 8


def _folder() -> str:
    base = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppLocalDataLocation)
    return os.path.join(base, "handtekeningen")


def list_saved():
    """[(pad, QImage)], nieuwste eerst."""
    folder = _folder()
    if not os.path.isdir(folder):
        return []
    paths = [os.path.join(folder, name) for name in os.listdir(folder) if name.endswith(".png")]
    paths.sort(key=os.path.getmtime, reverse=True)
    result = []
    for path in paths:
        image = QImage(path)
        if not image.isNull():
            result.append((path, image))
    return result


def save(image: QImage) -> str:
    """Bewaar een handtekening (dubbele worden niet opnieuw opgeslagen).
    Houdt maximaal MAX_SAVED handtekeningen; de oudste vervalt."""
    data = QByteArray()
    buffer = QBuffer(data)
    buffer.open(QIODevice.OpenModeFlag.WriteOnly)
    image.save(buffer, "PNG")
    buffer.close()
    raw = bytes(data)
    folder = _folder()
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, hashlib.sha1(raw).hexdigest()[:16] + ".png")
    with open(path, "wb") as handle:
        handle.write(raw)
    for old_path, _image in list_saved()[MAX_SAVED:]:
        delete(old_path)
    return path


def delete(path: str):
    try:
        os.remove(path)
    except OSError:
        pass
