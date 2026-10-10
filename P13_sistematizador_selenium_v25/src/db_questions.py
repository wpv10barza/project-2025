from __future__ import annotations
import pandas as pd
from typing import Optional
from .db import connect, read_db_config
from .sql import SQL_PREGUNTAS

def _pick_col(cols: list[str], preferred: list[str], contains_any: Optional[list[str]]=None) -> Optional[str]:
    lower_map = {c.lower(): c for c in cols}
    for p in preferred:
        if p.lower() in lower_map:
            return lower_map[p.lower()]
    if contains_any:
        for c in cols:
            cl = c.lower()
            if all(tok.lower() in cl for tok in contains_any):
                return c
    return None

def normalize_questions_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Devuelve un DF con columnas estándar:
    ID_PREGUNTA, PREGUNTA_CODIGO, PREGUNTA_DESCRIPCION, PREGUNTA_SITUACION_DESEADA
    """
    cols = list(df.columns)

    col_id = _pick_col(cols, ["ID_PREGUNTA", "PreguntaId", "Id"], contains_any=["pregunta", "id"]) or _pick_col(cols, ["Id"])
    col_codigo = _pick_col(cols, ["PREGUNTA_CODIGO", "PreguntaCodigo", "Codigo"], contains_any=["codigo"]) 
    col_desc = _pick_col(cols, ["PREGUNTA_DESCRIPCION", "PreguntaDescripcion", "Descripcion"], contains_any=["descripcion"])
    col_sit = _pick_col(cols, ["PREGUNTA_SITUACION_DESEADA", "PreguntaSituacionDeseada", "SituacionDeseada"], contains_any=["situacion","dese"])

    if col_codigo is None:
        raise ValueError(f"No pude detectar columna de código en VW_PREGUNTA_POR_EVALUACION. Columnas: {cols}")

    out = pd.DataFrame()
    out["ID_PREGUNTA"] = df[col_id] if col_id in df.columns else None
    out["PREGUNTA_CODIGO"] = df[col_codigo].astype(str).str.strip()
    out["PREGUNTA_DESCRIPCION"] = df[col_desc].astype(str) if col_desc in df.columns else ""
    out["PREGUNTA_SITUACION_DESEADA"] = df[col_sit].astype(str) if (col_sit and col_sit in df.columns) else ""

    # Filtrar filas sin código
    out = out[out["PREGUNTA_CODIGO"].notna() & (out["PREGUNTA_CODIGO"].astype(str).str.len() > 0)].copy()
    out = out.sort_values("PREGUNTA_CODIGO")
    return out

def load_questions(evaluacion_id: int, db_config_path) -> pd.DataFrame:
    cfg = read_db_config(db_config_path)
    conn = connect(cfg)
    try:
        df = pd.read_sql(SQL_PREGUNTAS, conn, params=[evaluacion_id])
    finally:
        conn.close()

    # Normaliza a columnas estándar para que el resto del pipeline no falle por nombres distintos
    return normalize_questions_df(df)
