# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
from typing import Iterable
from pypdf import PdfReader

def read_text_file(path: Path) -> str:
    return Path(path).read_text(encoding="utf-8", errors="ignore")

def read_pdf_text(path: Path) -> str:
    reader = PdfReader(str(path))
    out = []
    for page in reader.pages:
        out.append(page.extract_text() or "")
    return "\n".join(out)

def read_many_transcriptions(paths: Iterable[Path]) -> str:
    chunks = []
    for p in paths:
        if not p.exists():
            continue
        if p.suffix.lower() == ".pdf":
            chunks.append(read_pdf_text(p))
        else:
            chunks.append(read_text_file(p))
    return "\n\n".join([c.strip() for c in chunks if c and c.strip()])
