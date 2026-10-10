from __future__ import annotations
import os
import re
import csv
from pathlib import Path
from typing import List

from config import (
    OPENAI_MODEL, TEMPERATURE,
    PROMPT_TXT_OUT,
    RESULT_RAW_TXT, RESULT_TSV, RESULT_CSV, RESULT_XLSX,
)
from src.prompt_builder import write_prompt_files

def _extract_codeblock(text: str) -> str | None:
    m = re.search(r"```(?:tsv|csv)?\s*\n(.*?)\n```", text, flags=re.DOTALL | re.IGNORECASE)
    return m.group(1).strip() if m else None

def _guess_delimiter(sample: str) -> str:
    if "\t" in sample:
        return "\t"
    if ";" in sample and sample.count(";") >= sample.count(","):
        return ";"
    return ","

def _write_xlsx(rows: List[List[str]], out_path: Path) -> None:
    from openpyxl import Workbook  # lazy
    wb = Workbook()
    ws = wb.active
    ws.title = "resultado"
    for r in rows:
        ws.append(r)
    wb.save(out_path)

def run_gpt_and_save() -> None:
    try:
        from openai import OpenAI
    except Exception as e:
        raise RuntimeError(
            "No se pudo importar 'openai'.\n"
            "Solución: instala dependencias en tu venv con:\n"
            "  python -m pip install -r requirements.txt\n"
            f"Detalle: {e}"
        )

    if not os.getenv("OPENAI_API_KEY"):
        raise RuntimeError(
            "Falta OPENAI_API_KEY en el entorno.\n"
            "PowerShell:\n"
            "  $env:OPENAI_API_KEY='sk-...'\n"
            "Luego ejecuta:\n"
            "  python main.py run"
        )

    # Genera C8 y usa el prompt
    write_prompt_files()
    prompt = PROMPT_TXT_OUT.read_text(encoding="utf-8", errors="ignore")

    client = OpenAI()

    resp = client.responses.create(
        model=OPENAI_MODEL,
        temperature=TEMPERATURE,
        input=[
            {"role": "system", "content": "Responde únicamente con un único bloque TSV dentro de ```tsv ...``` y nada más."},
            {"role": "user", "content": prompt},
        ],
    )

    content = getattr(resp, "output_text", "") or ""
    RESULT_RAW_TXT.write_text(content, encoding="utf-8")

    table = _extract_codeblock(content) or content.strip()
    RESULT_TSV.write_text(table, encoding="utf-8")

    lines = [ln for ln in table.splitlines() if ln.strip()]
    if not lines:
        return

    delim = _guess_delimiter(lines[0])
    rows = [[cell.strip() for cell in row] for row in csv.reader(lines, delimiter=delim)]

    with RESULT_CSV.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerows(rows)

    _write_xlsx(rows, RESULT_XLSX)
