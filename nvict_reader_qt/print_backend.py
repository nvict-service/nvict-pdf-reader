# -*- coding: utf-8 -*-
"""Print-executielaag: printerlijst, paginabereik-parsing en de printjob.

Poort van NVict_Reader.py's execute_print (regel 3615-3891), maar via
QPrinter/QPainter i.p.v. win32print/win32ui/PIL+ImageWin - zie het fase-2
plan voor de onderbouwing (QPrinterInfo verwijdert de pywin32-afhankelijkheid
volledig, en QPageSetupDialog vervangt win32print.DocumentProperties).
"""

from dataclasses import dataclass

from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QImage, QPainter, QTransform
from PySide6.QtPrintSupport import QPrinter, QPrinterInfo

from .document import get_fitz

# Renderresolutie: 300 dpi is scherp voor tekst en houdt een A4 onder ~30 MB
# geheugen. Hogere printer-dpi's schaalt de driver zelf op (geen kwaliteitsverlies
# t.o.v. 300 dpi, want we schalen niet zelf nog eens in Python).
MAX_PRINT_ZOOM = 300 / 72


@dataclass
class PrintOptions:
    printer_name: str
    pages: list  # 0-indexed paginanummers
    copies: int = 1
    fit_to_page: bool = True
    duplex: bool = False
    manual_duplex: bool = False  # voorkanten, gebruiker draait de stapel om, achterkanten
    color_mode: str = "kleur"  # "kleur" / "zwart_wit"
    rotation: int = 0  # 0/90/180/270
    orientation: str = "staand"  # "staand" / "liggend"


def list_printer_names():
    return [p.printerName() for p in QPrinterInfo.availablePrinters()]


def default_printer_name():
    printer = QPrinterInfo.defaultPrinter()
    return printer.printerName() if not printer.isNull() else ""


def parse_page_range(page_string, total_pages):
    """Parse '1,3,5' of '1-5,7' naar een gesorteerde lijst 0-indexed paginanummers.

    Verbatim poort van NVict_Reader.py:3573-3613. Geeft None terug bij een
    ongeldige/out-of-range invoer.
    """
    pages = set()
    try:
        page_string = page_string.replace(" ", "")
        for part in page_string.split(","):
            if "-" in part:
                start, end = part.split("-")
                start, end = int(start), int(end)
                if start < 1 or end > total_pages or start > end:
                    return None
                for page_num in range(start, end + 1):
                    pages.add(page_num - 1)
            else:
                page_num = int(part)
                if page_num < 1 or page_num > total_pages:
                    return None
                pages.add(page_num - 1)
        return sorted(pages)
    except (ValueError, AttributeError):
        return None


# Printers zonder gewone vellen papier (virtueel of labels/tape): handmatig
# dubbelzijdig is daar zinloos.
_VIRTUAL_PRINTER_HINTS = (
    "pdf", "onenote", "xps", "fax",
    "label", "brother ql", "brother pt", "dymo", "zebra",
)


def is_virtual_printer(name: str) -> bool:
    lowered = (name or "").lower()
    return any(hint in lowered for hint in _VIRTUAL_PRINTER_HINTS)


def manual_duplex_rounds(pages, copies, reverse_backs=False):
    """Verdeel de pagina's over vellen voor handmatig dubbelzijdig printen.

    Geeft (voorkanten, achterkanten) terug: twee even lange lijsten met per
    vel een paginanummer, of None voor een lege achterkant (laatste vel bij
    een oneven aantal pagina's). Die lege pagina wordt wél "geprint", zodat
    vel en achterkant gelijk blijven lopen. Met `reverse_backs` komen de
    achterkanten in omgekeerde volgorde - nodig bij printers die de vellen
    met de bedrukte kant naar boven uitvoeren.
    """
    fronts, backs = [], []
    for _copy in range(max(copies, 1)):
        fronts.extend(pages[0::2])
        copy_backs = list(pages[1::2])
        if len(pages) % 2:
            copy_backs.append(None)
        backs.extend(copy_backs)
    if reverse_backs:
        backs.reverse()
    return fronts, backs


def render_page_for_print(pdf_document, page_num, zoom, rotation=0, color_mode="kleur"):
    """Render één pagina op printresolutie, met rotatie/kleurmodus toegepast."""
    fitz = get_fitz()
    page = pdf_document[page_num]
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888).copy()

    if rotation in (90, 180, 270):
        image = image.transformed(QTransform().rotate(rotation), Qt.TransformationMode.SmoothTransformation)

    if color_mode == "zwart_wit":
        image = image.convertToFormat(QImage.Format.Format_Grayscale8).convertToFormat(QImage.Format.Format_RGB888)

    return image


def run_print_job(printer: QPrinter, pdf_document, options: PrintOptions, on_progress=None, should_cancel=None,
                  sequence=None):
    """Voer de printjob uit. `on_progress(current, total)` en `should_cancel()`
    zijn optionele callbacks voor een voortgangsdialoog met annuleren.

    `sequence` is een expliciete lijst paginanummers (None = lege pagina),
    voor handmatig dubbelzijdig; anders options.pages × options.copies.

    Geeft True terug bij volledige afronding, False als geannuleerd.
    """
    if sequence is None:
        sequence = list(options.pages) * max(options.copies, 1)
    total_units = len(sequence)
    printer_rect = printer.pageRect(QPrinter.Unit.DevicePixel)
    printer_width, printer_height = printer_rect.width(), printer_rect.height()
    dpi_scale = max(printer.logicalDpiX(), printer.logicalDpiY()) / 72
    zoom = min(dpi_scale, MAX_PRINT_ZOOM)
    # Printerpixels per renderpixel, voor de werkelijke afmeting op 100%.
    device_per_image_px = dpi_scale / zoom

    painter = QPainter(printer)
    painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
    try:
        done = 0
        first_page = True
        for page_num in sequence:
            if should_cancel and should_cancel():
                return False
            if not first_page:
                printer.newPage()
            first_page = False

            if page_num is None:  # lege achterkant
                done += 1
                if on_progress:
                    on_progress(done, total_units)
                continue

            try:
                image = render_page_for_print(pdf_document, page_num, zoom, options.rotation, options.color_mode)
            except Exception as page_error:
                print(f"Fout bij printen pagina {page_num + 1}: {page_error}")
                done += 1
                if on_progress:
                    on_progress(done, total_units)
                continue

            img_width, img_height = image.width(), image.height()
            aspect_ratio = img_width / img_height if img_height else 1.0

            if options.fit_to_page:
                printer_aspect = printer_width / printer_height if printer_height else 1.0
                if aspect_ratio > printer_aspect:
                    print_width = printer_width
                    print_height = int(printer_width / aspect_ratio)
                else:
                    print_height = printer_height
                    print_width = int(printer_height * aspect_ratio)
            else:
                # Fysieke afmeting op 100% - alleen naar beneden afschalen
                # als hij niet past (nooit uitrekken).
                print_width = img_width * device_per_image_px
                print_height = img_height * device_per_image_px
                if print_width > printer_width or print_height > printer_height:
                    scale = min(printer_width / print_width, printer_height / print_height)
                    print_width = int(print_width * scale)
                    print_height = int(print_height * scale)

            x = (printer_width - print_width) // 2
            y = (printer_height - print_height) // 2
            # Het beeld ongeschaald in een doelrechthoek tekenen: de driver
            # schaalt het dan zelf, zonder extra resampling-verlies.
            painter.drawImage(QRectF(x, y, print_width, print_height), image)

            done += 1
            if on_progress:
                on_progress(done, total_units)
    finally:
        painter.end()

    return True
