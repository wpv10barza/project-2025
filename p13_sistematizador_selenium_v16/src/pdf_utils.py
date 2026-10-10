# -*- coding: utf-8 -*-
from pathlib import Path
from pypdf import PdfReader

def read_pdf_text(path: Path) -> str:
    reader = PdfReader(str(path))
    parts = []
    for page in reader.pages:
        parts.append(page.extract_text() or "")
    return "\n".join(parts).strip()
