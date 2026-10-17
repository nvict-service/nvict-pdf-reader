# -*- coding: utf-8 -*-
"""Een PDF-pagina renderen naar een QPixmap die scherp is op het scherm.

Bij een Windows-schermschaal boven 100% (125%, 150%, ...) heeft Qt per
logische pixel meer dan één echte schermpixel. Renderden we op alleen
`zoom`, dan rekte Qt die bitmap op en werd de tekst wazig (klantmelding:
"net of de resolutie niet juist is"). We renderen daarom op
`zoom × device_pixel_ratio` en zetten die ratio op de pixmap, zodat hij
in logische pixels even groot blijft maar elke schermpixel vult.
"""

from PySide6.QtGui import QImage, QPixmap

from .document import get_fitz

# Leesmodus (Beeld-menu): pagina's donker met lichte tekst. Module-breed,
# zodat weergave, miniaturen en volledig scherm hetzelfde tonen; het
# welkomscherm geeft reading=False mee (bestandsvoorbeelden).
reading_mode = False
READING_BACKGROUND = 0x202124
READING_TEXT = 0xE8E8E8


def render_page(page, zoom, device_pixel_ratio=1.0, reading=None) -> QPixmap:
    fitz = get_fitz()
    scale = zoom * device_pixel_ratio
    if reading is None:
        reading = reading_mode
    matrix = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=matrix)
    if reading:
        original = None
        # Omkeren en dan "zwart" -> donkergrijs, "wit" -> lichtgrijs: zachter
        # voor de ogen dan puur zwart/wit.
        pix.invert_irect()
        pix.tint_with(READING_BACKGROUND, READING_TEXT)
        # Foto's niet als negatief tonen: die vlakken uit de gewone render
        # terugzetten.
        try:
            image_boxes = [fitz.Rect(info["bbox"]) for info in page.get_image_info()]
        except Exception:
            image_boxes = []
        for box in image_boxes:
            if box.is_empty or box.is_infinite:
                continue
            if original is None:
                original = page.get_pixmap(matrix=matrix)
            pix.copy(original, (box * matrix).irect & pix.irect)
    fmt = QImage.Format.Format_RGBA8888 if pix.alpha else QImage.Format.Format_RGB888
    image = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
    pixmap = QPixmap.fromImage(image.copy())
    pixmap.setDevicePixelRatio(device_pixel_ratio)
    return pixmap
