from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple, Optional

import pandas as pd


@dataclass
class MatrixSpec:
    interviewer_speakers: List[str]
    informant_speakers: List[str]

    @property
    def columns(self) -> List[str]:
        interviewer_cols = [f"CONVERSACION_ENTREVISTADOR_{i+1}" for i in range(len(self.interviewer_speakers))]
        informant_cols = [f"CONVERSACION_INFORMANTE_{i+1}" for i in range(len(self.informant_speakers))]
        return [
            "PREGUNTA_CODIGO",
            *interviewer_cols,
            *informant_cols,
            "PREGUNTA_DESCRIPCION",
            "PREGUNTA_SITUACION_DESEADA",
            "ANALISIS_DE_COBERTURA",
        ]


def parse_tsv_row(line: str, expected_cols: List[str]) -> Dict[str, str]:
    parts = line.rstrip("\n").split("\t")
    if len(parts) < len(expected_cols):
        parts = parts + [""] * (len(expected_cols) - len(parts))
    elif len(parts) > len(expected_cols):
        # Une extras al final (última columna)
        parts = parts[: len(expected_cols) - 1] + ["\t".join(parts[len(expected_cols) - 1 :])]

    return {col: parts[i] if i < len(parts) else "" for i, col in enumerate(expected_cols)}


def rows_to_dataframe(rows: List[Dict[str, str]], columns: List[str]) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    # Asegura orden de columnas
    for c in columns:
        if c not in df.columns:
            df[c] = ""
    df = df[columns]
    return df


def save_outputs(df: pd.DataFrame, output_dir: Path, base_name: str = "systematization_matrix") -> Tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    csv_path = output_dir / f"{base_name}.csv"
    xlsx_path = output_dir / f"{base_name}.xlsx"

    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    df.to_excel(xlsx_path, index=False)

    return csv_path, xlsx_path


def save_rows_tsv(lines: List[str], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")
