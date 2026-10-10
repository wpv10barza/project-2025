from __future__ import annotations
from pathlib import Path
from pypdf import PdfReader

def read_pdf_text(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    parts: list[str] = []
    for i, page in enumerate(reader.pages):
        txt = page.extract_text() or ""
        txt = txt.replace("\r\n", "\n").replace("\r", "\n")
        parts.append(f"\n\n--- PAGE {i+1} ---\n{txt}")
    return "\n".join(parts).strip()
