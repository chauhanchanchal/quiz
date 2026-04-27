"""
ocr_extract.py
Local text extractor for handwritten/printed PDFs and images.
Uses EasyOCR — runs fully offline, no API key, no credit card.
"""

import os
import io
from pathlib import Path
import easyocr
from pdf2image import convert_from_path
from PIL import Image
import numpy as np

POPPLER_PATH = r"C:\poppler\Library\bin\pdftoppm.exe\poppler-25.12.0\Library\bin"
LANGUAGES    = ["en"]
USE_GPU      = True

_reader = None
def _get_reader():
    global _reader
    if _reader is None:
        print("Loading EasyOCR model (first run downloads ~64 MB)...")
        _reader = easyocr.Reader(LANGUAGES, gpu=USE_GPU)
        print("EasyOCR ready.")
    return _reader


def _extract_from_pil(pil_image: Image.Image) -> str:
    reader = _get_reader()
    img_array = np.array(pil_image.convert("RGB"))
    results = reader.readtext(img_array, detail=1, paragraph=False)

    def sort_key(r):
        bbox = r[0]
        ys = [pt[1] for pt in bbox]
        xs = [pt[0] for pt in bbox]
        return (min(ys), min(xs))

    results.sort(key=sort_key)

    lines = []
    current_line = []
    last_y = None
    for bbox, text, conf in results:
        y_center = sum(pt[1] for pt in bbox) / 4
        if last_y is None or abs(y_center - last_y) < 20:
            current_line.append((bbox, text, conf))
        else:
            lines.append(current_line)
            current_line = [(bbox, text, conf)]
        last_y = y_center
    if current_line:
        lines.append(current_line)

    line_texts = []
    for line in lines:
        line.sort(key=lambda r: min(pt[0] for pt in r[0]))
        line_texts.append(" ".join(item[1] for item in line))

    return "\n".join(line_texts)


def extract_text(file_path: str) -> str:
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    suffix = file_path.suffix.lower()

    if suffix in {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".tif"}:
        image = Image.open(file_path)
        return _extract_from_pil(image)

    if suffix == ".pdf":
        try:
            pages = convert_from_path(
                str(file_path),
                dpi=200,
                poppler_path=POPPLER_PATH if os.path.isdir(POPPLER_PATH) else None,
            )
        except Exception as e:
            raise RuntimeError(
                f"Failed to convert PDF. Is poppler installed at {POPPLER_PATH}? "
                f"Error: {e}"
            )

        all_text = []
        for i, page in enumerate(pages, start=1):
            try:
                text = _extract_from_pil(page)
            except Exception as e:
                text = f"[Page {i} OCR failed: {e}]"
            all_text.append(f"--- Page {i} ---\n{text}")
        return "\n\n".join(all_text)

    raise ValueError(f"Unsupported file type: {suffix}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python ocr_extract.py <path-to-pdf-or-image>")
        sys.exit(1)
    text = extract_text(sys.argv[1])
    print("=" * 60)
    print(text)
    print("=" * 60)
    print(f"\nExtracted {len(text)} characters.")