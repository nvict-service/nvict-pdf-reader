# -*- coding: utf-8 -*-
"""Print-executielaag: printerlijst, paginabereik-parsing en de printjob.

Poort van NVict_Reader.py's execute_print (regel 3615-3891), maar via
QPrinter/QPainter i.p.v. win32print/win32ui/PIL+ImageWin - zie het fase-2
plan voor de onderbouwing (QPrinterInfo verwijdert de pywin32-afhankelijkheid
volledig, en QPageSetupDialog vervangt win32print.DocumentProperties).
"""

from dataclasses import dataclass

from PySide6.QtCore import Qt
from PySide6.QtGui import QImage, QPainter, QTransform
from PySide6.QtPrintSupport import QPrinter, QPrinterInfo

from .document import get_fitz

MAX_PRINT_ZOOM = 4.0  # zelfde cap als NVict_Reader.py:3773


@dataclass
class PrintOptions:
    printer_name: str
    pages: list  # 0-indexed paginanummers
    copies: int = 1
    fit_to_page: bool = True
    duplex: bool = False
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


def run_print_job(printer: QPrinter, pdf_document, options: PrintOptions, on_progress=None, should_cancel=None):
    """Voer de printjob uit. `on_progress(current, total)` en `should_cancel()`
    zijn optionele callbacks voor een voortgangsdialoog met annuleren.

    Geeft True terug bij volledige afronding, False als geannuleerd.
    """
    total_units = len(options.pages) * max(options.copies, 1)
    printer_rect = printer.pageRect(QPrinter.Unit.DevicePixel)
    printer_width, printer_height = printer_rect.width(), printer_rect.height()
    dpi_scale = max(printer.logicalDpiX(), printer.logicalDpiY()) / 72
    zoom = min(dpi_scale, MAX_PRINT_ZOOM)

    painter = QPainter(printer)
    try:
        done = 0
        first_page = True
        for _copy in range(max(options.copies, 1)):
            for page_num in options.pages:
                if should_cancel and should_cancel():
                    return False
                if not first_page:
                    printer.newPage()
                first_page = False

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
                    # Afbeelding is al op printresolutie gerenderd, dus dit IS
                    # de fysieke afmeting op 100% - alleen naar beneden
                    # afschalen als hij niet past (nooit uitrekken).
                    print_width, print_height = img_width, img_height
                    if print_width > printer_width or print_height > printer_height:
                        scale = min(printer_width / print_width, printer_height / print_height)
                        print_width = int(print_width * scale)
                        print_height = int(print_height * scale)

                x = (printer_width - print_width) // 2
                y = (printer_height - print_height) // 2
                scaled = image.scaled(
                    int(print_width), int(print_height),
                    Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation,
                )
                painter.drawImage(int(x), int(y), scaled)

                done += 1
                if on_progress:
                    on_progress(done, total_units)
    finally:
        painter.end()

    return True
