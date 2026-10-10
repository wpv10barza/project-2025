from __future__ import annotations
from pathlib import Path

def read_pdf_text(path: Path) -> str:
    try:
        from PyPDF2 import PdfReader
    except Exception as e:
        raise RuntimeError("Falta PyPDF2. Instala con: python -m pip install PyPDF2") from e

    p = Path(path)
    reader = PdfReader(str(p))
    parts = []
    for page in reader.pages:
        txt = page.extract_text() or ""
        parts.append(txt)
    return "\n".join(parts).strip()
