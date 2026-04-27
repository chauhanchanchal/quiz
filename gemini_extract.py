"""
gemini_extract.py
Calls Google Gemini directly via HTTPS — no SDK, no pydantic, no compiled DLLs.
Uses the same `requests` library your app already uses for Groq.

KEY FIX: When Gemini fails on a page, this version RAISES an exception
instead of returning the error message as transcribed text. This way,
app.py's EasyOCR fallback actually triggers correctly.
"""

import os
import io
import base64
import json
import time
from pathlib import Path
import requests
from PIL import Image
from pdf2image import convert_from_path

# --- CONFIG -------------------------------------------------------------------
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
# Cross-platform: explicit path on Windows, system default on Linux (HF Spaces)
import platform
if platform.system() == "Windows":
    POPPLER_PATH = r"C:\poppler\Library\bin\pdftoppm.exe\poppler-25.12.0\Library\bin"
else:
    POPPLER_PATH = ""
MODEL_NAME = "gemini-flash-latest"

GEMINI_URL = (
    f"https://generativelanguage.googleapis.com/v1beta/models/"
    f"{MODEL_NAME}:generateContent"
)

# --- TRANSCRIPTION PROMPT -----------------------------------------------------
TRANSCRIBE_PROMPT = """You are an expert at reading handwritten exam answer sheets.

Look at this handwritten page carefully and transcribe ALL the text you can see.

Rules:
1. Preserve the question structure (Q1, Q1a, Q2, etc.) exactly as written.
2. Preserve line breaks and paragraph structure.
3. For code or math, transcribe symbols accurately (=, +, -, *, /, parentheses, etc.).
4. If a word is unclear, give your best guess in [brackets] like [unclear: maybe "function"].
5. Do NOT summarize. Do NOT add commentary. Output ONLY the transcribed text.
6. If the page is blank or has no text, output exactly: [BLANK PAGE]
7. Preserve the student's spelling and grammar — do not correct mistakes.

Begin transcription now:"""


# --- HELPERS ------------------------------------------------------------------
def _pil_to_base64(pil_image: Image.Image) -> str:
    """Convert a PIL image to base64-encoded PNG bytes."""
    buf = io.BytesIO()
    pil_image.convert("RGB").save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode("utf-8")


def _transcribe_image(pil_image: Image.Image, page_num: int = 1) -> str:
    """Send a PIL image to Gemini via HTTP and return transcribed text.

    Raises RuntimeError on any failure — caller is responsible for fallback.
    """
    if not GEMINI_API_KEY:
        raise RuntimeError(
            "GEMINI_API_KEY environment variable not set. "
            "Set it with: $env:GEMINI_API_KEY = 'your_key'"
        )

    image_b64 = _pil_to_base64(pil_image)

    payload = {
        "contents": [{
            "parts": [
                {"text": TRANSCRIBE_PROMPT},
                {
                    "inline_data": {
                        "mime_type": "image/png",
                        "data": image_b64,
                    }
                },
            ]
        }],
        "generationConfig": {
            "temperature": 0.1,
            "maxOutputTokens": 4096,
        },
    }

    try:
        response = requests.post(
            GEMINI_URL,
            params={"key": GEMINI_API_KEY},
            headers={"Content-Type": "application/json"},
            json=payload,
            timeout=60,
        )
    except requests.RequestException as e:
        raise RuntimeError(f"Network error on page {page_num}: {e}")

    if response.status_code != 200:
        # Truncate to keep the error message manageable
        err_text = response.text[:300] if response.text else "(no body)"
        raise RuntimeError(
            f"Gemini API error on page {page_num} "
            f"(HTTP {response.status_code}): {err_text}"
        )

    try:
        data = response.json()
        text = data["candidates"][0]["content"]["parts"][0]["text"]
        return text.strip()
    except (KeyError, IndexError, json.JSONDecodeError) as e:
        raise RuntimeError(
            f"Could not parse Gemini response on page {page_num}: {e}. "
            f"Raw response: {response.text[:300]}"
        )


# --- PUBLIC API ---------------------------------------------------------------
def extract_text(file_path: str) -> str:
    """Extract text from PDF or image using Gemini vision.

    Raises RuntimeError if Gemini fails (so app.py's EasyOCR fallback triggers).
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    suffix = file_path.suffix.lower()

    # Single image — fail loudly if Gemini fails
    if suffix in {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".tif"}:
        image = Image.open(file_path)
        return _transcribe_image(image, 1)

    # PDF — process page by page; tolerate partial failures, fail if too many
    if suffix == ".pdf":
        try:
            pages = convert_from_path(
                str(file_path),
                dpi=200,
                poppler_path=POPPLER_PATH if os.path.isdir(POPPLER_PATH) else None,
            )
        except Exception as e:
            raise RuntimeError(f"PDF to image conversion failed: {e}")

        all_text = []
        failed_pages = 0
        first_error = None

        for i, page in enumerate(pages, start=1):
            print(f"  Gemini transcribing page {i}/{len(pages)}...")
            if i > 1:
                time.sleep(4.5)  # stay under 15 req/min on free tier
            try:
                text = _transcribe_image(page, i)
                all_text.append(f"--- Page {i} ---\n{text}")
            except Exception as e:
                failed_pages += 1
                if first_error is None:
                    first_error = str(e)
                print(f"  Page {i} FAILED: {e}")

                # If page 1 fails, abort immediately so EasyOCR fallback runs.
                # Better to OCR the whole doc with EasyOCR than mix Gemini + nothing.
                if i == 1:
                    raise RuntimeError(
                        f"Gemini failed on first page — aborting so fallback can run. "
                        f"First error: {first_error}"
                    )

                # Otherwise note it but keep trying remaining pages
                all_text.append(f"--- Page {i} ---\n[Transcription unavailable]")

        # If most pages failed, raise so the caller can fall back to EasyOCR
        if failed_pages > len(pages) / 2:
            raise RuntimeError(
                f"Gemini failed on {failed_pages}/{len(pages)} pages "
                f"— quota likely exhausted. First error: {first_error}"
            )

        return "\n\n".join(all_text)

    raise ValueError(f"Unsupported file type: {suffix}")


# --- CLI ----------------------------------------------------------------------
if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python gemini_extract.py <path-to-pdf-or-image>")
        sys.exit(1)
    try:
        text = extract_text(sys.argv[1])
        print("=" * 60)
        print(text)
        print("=" * 60)
        print(f"\nExtracted {len(text)} characters.")
    except Exception as e:
        print(f"FAILED: {e}")
        sys.exit(1)
