from __future__ import annotations

from pathlib import Path
from typing import Optional
from PyPDF2 import PdfReader


def read_pdf_text(pdf_path: Path) -> str:
    """Extrae texto de un PDF (sin OCR)."""
    reader = PdfReader(str(pdf_path))
    parts = []
    for page in reader.pages:
        try:
            parts.append(page.extract_text() or "")
        except Exception:
            parts.append("")
    # Separa páginas con salto de línea para evitar pegar palabras
    return "\n".join(parts)


def find_first_pdf_by_keywords(folder: Path, keywords: list[str]) -> Optional[Path]:
    """Busca el primer PDF en `folder` cuyo nombre contenga alguno de los keywords (case-insensitive)."""
    pdfs = sorted(folder.glob("*.pdf"))
    lower_keywords = [k.lower() for k in keywords]
    for p in pdfs:
        name = p.name.lower()
        if any(k in name for k in lower_keywords):
            return p
    return None
