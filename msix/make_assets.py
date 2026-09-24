#!/usr/bin/env python3
"""Genereer de MSIX/Store-afbeeldingen in msix/Assets uit logo.png en
pdf_file_icon.ico.

Alleen opnieuw draaien als het logo verandert; de PNG's staan in git.

    python msix/make_assets.py

De bestandsnamen volgen de MRT-qualifiers (scale-/targetsize-/altform-);
_nvict_build maakt bij het verpakken een resources.pri aan zodat Windows per
schermschaal en per plek (taakbalk, Start, Verkenner) de juiste kiest.
"""
from pathlib import Path

from PIL import Image

HERE = Path(__file__).resolve().parent
APP_DIR = HERE.parent
OUT = HERE / "Assets"

LOGO = APP_DIR / "logo.png"
PDF_ICON = APP_DIR / "pdf_file_icon.ico"


def _largest_frame(path: Path) -> Image.Image:
    img = Image.open(path)
    sizes = img.info.get("sizes")
    if sizes:
        img.size = max(sizes)  # ICO: kies het grootste frame
    return img.convert("RGBA")


def _square(src: Image.Image, size: int, padding: float = 0.0) -> Image.Image:
    """Schaal src in een transparant vierkant van size x size, met optionele
    marge (fractie van de rand) zodat tegels niet tegen de rand aan lopen."""
    inner = max(1, round(size * (1 - 2 * padding)))
    im = src.copy()
    im.thumbnail((inner, inner), Image.LANCZOS)
    canvas = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    canvas.paste(im, ((size - im.width) // 2, (size - im.height) // 2), im)
    return canvas


def _wide(src: Image.Image, width: int, height: int) -> Image.Image:
    inner = round(height * 0.8)
    im = src.copy()
    im.thumbnail((inner, inner), Image.LANCZOS)
    canvas = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    canvas.paste(im, ((width - im.width) // 2, (height - im.height) // 2), im)
    return canvas


def main() -> None:
    OUT.mkdir(exist_ok=True)
    for old in OUT.glob("*.png"):
        old.unlink()

    logo = Image.open(LOGO).convert("RGBA")
    pdf = _largest_frame(PDF_ICON)

    def save(img: Image.Image, name: str) -> None:
        img.save(OUT / name, optimize=True)

    for scale in (100, 200, 400):
        f = scale / 100
        save(_square(logo, round(44 * f)), f"Square44x44Logo.scale-{scale}.png")
        save(_square(logo, round(150 * f), padding=0.15), f"Square150x150Logo.scale-{scale}.png")
        save(_square(logo, round(50 * f)), f"StoreLogo.scale-{scale}.png")
        save(_square(pdf, round(44 * f)), f"PdfFile.scale-{scale}.png")
    for scale in (100, 200):
        f = scale / 100
        save(_wide(logo, round(310 * f), round(150 * f)), f"Wide310x150Logo.scale-{scale}.png")

    for size in (16, 24, 32, 48, 256):
        save(_square(logo, size), f"Square44x44Logo.targetsize-{size}.png")
        save(_square(logo, size), f"Square44x44Logo.targetsize-{size}_altform-unplated.png")
        save(_square(pdf, size), f"PdfFile.targetsize-{size}.png")

    print(f"{len(list(OUT.glob('*.png')))} afbeeldingen geschreven naar {OUT}")


if __name__ == "__main__":
    main()
