# -*- coding: utf-8 -*-
from __future__ import annotations

import pandas as pd

def print_excel_vs_db_report(df_excel: pd.DataFrame, df_q: pd.DataFrame, col_preg_codigo: str = "PREGUNTA_CODIGO") -> dict:
    q_codes = set(df_q["PREGUNTA_CODIGO"].astype(str).str.strip())
    x_codes = set(df_excel[col_preg_codigo].astype(str).str.strip())

    match = sorted(q_codes.intersection(x_codes))
    extra = sorted(x_codes - q_codes)
    missing = sorted(q_codes - x_codes)

    stats = {
        "preguntas_bd": len(q_codes),
        "codigos_excel": len(x_codes),
        "codigos_match": len(match),
        "extra_excel": len(extra),
        "faltan_en_excel": len(missing),
    }
    return stats
