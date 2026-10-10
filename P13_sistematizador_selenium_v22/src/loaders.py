from __future__ import annotations
from pathlib import Path
import pandas as pd
from typing import Optional
from .db import connect, read_db_config
from .db_questions import load_questions
from .excel_reader import read_excel_matrix
from .mapping import map_row_to_respuesta
from .repo import get_aplicacion_id
from config import DEFAULT_USUARIO_REGISTRO_ID, DEFAULT_TABLE

def load_excel_to_sql(
    evaluacion_id: int,
    informante_id: int,
    excel_path: Path,
    db_config_path: Path,
    table: str = DEFAULT_TABLE,
    usuario_registro_id: int = DEFAULT_USUARIO_REGISTRO_ID,
) -> int:
    """
    Carga la matriz desde Excel y hace MERGE en dbo.Respuesta usando nombres reales:
      AplicacionId, PreguntaId, SituacionActual, Notas, RespuestaLiteral, ...
    """
    cfg = read_db_config(db_config_path)
    aplicacion_id = get_aplicacion_id(evaluacion_id, informante_id, db_config_path)
    print(f"[OK] AplicacionId: {aplicacion_id}")

    # Preguntas desde la vista (normalizadas sin depender de nombres exactos)
    df_preg = load_questions(evaluacion_id=evaluacion_id, db_config_path=db_config_path)
    code_to_id = {str(r["PREGUNTA_CODIGO"]).strip(): int(r["ID_PREGUNTA"]) for _, r in df_preg.iterrows() if pd.notna(r["ID_PREGUNTA"])}

    df_xls = read_excel_matrix(Path(excel_path))

    conn = connect(cfg)
    rows_ok = 0
    try:
        cur = conn.cursor()

        merge_sql = f"""
MERGE {table} AS tgt
USING (SELECT ? AS AplicacionId, ? AS PreguntaId) AS src
ON (tgt.AplicacionId = src.AplicacionId AND tgt.PreguntaId = src.PreguntaId)
WHEN MATCHED THEN
  UPDATE SET
    tgt.SituacionActual = ?,
    tgt.Notas = ?,
    tgt.RespuestaLiteral = ?,
    tgt.UsuarioActualizacionId = ?,
    tgt.FechaActualizacion = GETDATE()
WHEN NOT MATCHED THEN
  INSERT (AplicacionId, PreguntaId, SituacionActual, Notas, RespuestaLiteral, UsuarioRegistroId, FechaRegistro)
  VALUES (?, ?, ?, ?, ?, ?, GETDATE());
"""
        for _, row in df_xls.iterrows():
            codigo = str(row.get("PREGUNTA_CODIGO", "")).strip()
            if not codigo:
                continue
            pregunta_id = code_to_id.get(codigo)
            if not pregunta_id:
                # código no existe para esa evaluación
                continue

            mapped = map_row_to_respuesta(row=row, pregunta_id=pregunta_id, aplicacion_id=aplicacion_id)

            # Reglas mínimas: SituacionActual y Notas NO NULL en dbo.Respuesta (según tu imagen)
            if not mapped["SituacionActual"] or not mapped["Notas"]:
                continue

            cur.execute(
                merge_sql,
                # src key
                aplicacion_id, pregunta_id,
                # matched update values
                mapped["SituacionActual"], mapped["Notas"], mapped["RespuestaLiteral"],
                usuario_registro_id,
                # insert values
                aplicacion_id, pregunta_id,
                mapped["SituacionActual"], mapped["Notas"], mapped["RespuestaLiteral"],
                usuario_registro_id,
            )
            rows_ok += 1

        conn.commit()
        print(f"[OK] Filas procesadas: {rows_ok}")
        return rows_ok
    finally:
        conn.close()
