# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import List
from openpyxl import Workbook

def extract_codeblock(text: str) -> str | None:
    m = re.search(r"```(?:tsv|csv)?\s*\n(.*?)\n```", text, flags=re.DOTALL | re.IGNORECASE)
    return m.group(1).strip() if m else None

def guess_delimiter(sample: str) -> str:
    if "\t" in sample:
        return "\t"
    if ";" in sample and sample.count(";") >= sample.count(","):
        return ";"
    return ","

def write_xlsx(rows: List[List[str]], out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "resultado"
    for r in rows:
        ws.append(r)
    wb.save(out_path)

def write_table_files(table_text: str, tsv_path: Path, csv_path: Path, xlsx_path: Path) -> None:
    tsv_path.write_text(table_text, encoding="utf-8")
    lines = [ln for ln in table_text.splitlines() if ln.strip()]
    if not lines:
        return
    delim = guess_delimiter(lines[0])
    rows = [[cell for cell in row] for row in csv.reader(lines, delimiter=delim)]
    with csv_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerows(rows)
    write_xlsx(rows, xlsx_path)
