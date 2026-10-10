# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
from pathlib import Path

def docx_to_text(docx_path: Path) -> str:
    # Lazy import so preview/load can work even if python-docx isn't installed yet
    from docx import Document  # type: ignore

    doc = Document(str(docx_path))
    parts = []
    for p in doc.paragraphs:
        t = (p.text or "").rstrip()
        if t:
            parts.append(t)
    return "\n".join(parts).strip()

def main() -> int:
    ap = argparse.ArgumentParser(description="Convierte .docx a Transcripcion.txt (UTF-8).")
    ap.add_argument("--docx", required=True, help="Ruta al .docx de transcripción")
    ap.add_argument("--out", required=True, help="Ruta de salida .txt (por ejemplo ...\Transcripcion.txt)")
    args = ap.parse_args()

    docx_path = Path(args.docx)
    out_path = Path(args.out)

    if not docx_path.exists():
        raise FileNotFoundError(f"No existe DOCX: {docx_path}")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    text = docx_to_text(docx_path)
    out_path.write_text(text, encoding="utf-8")

    print(f"[OK] TXT generado: {out_path}")
    print(f"     Caracteres: {len(text)}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
