# -*- coding: utf-8 -*-
from pathlib import Path
from src.pdf_utils import read_pdf_text

def read_transcription(path: Path) -> str:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"No existe transcripción en: {p}")

    ext = p.suffix.lower()
    if ext == ".txt":
        return p.read_text(encoding="utf-8", errors="ignore").strip()
    if ext == ".pdf":
        return read_pdf_text(p)
    if ext == ".docx":
        from docx import Document
        doc = Document(str(p))
        return "\n".join([para.text for para in doc.paragraphs if para.text is not None]).strip()

    raise ValueError(f"Formato de transcripción no soportado: {ext} (usa .txt/.pdf/.docx)")
