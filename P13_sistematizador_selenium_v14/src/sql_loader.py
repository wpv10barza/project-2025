# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
import pandas as pd

from config import DB_CONFIG_FILE, DB_CONFIG_FILE_FALLBACK, RESULT_XLSX
from src.db import read_db_config, connect
from src.db_questions import fetch_questions
from src.diagnostics import print_excel_vs_db_report

SQL_APLICACION = """
SELECT TOP (1) Id
FROM dbo.Aplicacion
WHERE EvaluacionId = ? AND InformanteId = ?
ORDER BY Id DESC;
"""

# Ajusta aquí el nombre real de la tabla destino (según tu BD):
TABLA_DESTINO = "dbo.Respuesta"  # o dbo.Respuesta_Prueba_Plaud

def _obtener_aplicacion_id(conn, evaluacion_id: int, informante_id: int) -> int:
    df = pd.read_sql(SQL_APLICACION, conn, params=[evaluacion_id, informante_id])
    if df.empty:
        raise RuntimeError(f"No existe Aplicacion para EvaluacionId={evaluacion_id} e InformanteId={informante_id}.")
    return int(df.iloc[0]["Id"])

def load_to_sql(evaluacion_id: int, informante_id: int, excel_path: str | None, db_config_path: str | None) -> dict:
    xls = Path(excel_path) if excel_path else RESULT_XLSX
    if not xls.exists():
        raise FileNotFoundError(f"No existe Excel a cargar: {xls}")

    cfg_path = Path(db_config_path) if db_config_path else (DB_CONFIG_FILE if DB_CONFIG_FILE.exists() else DB_CONFIG_FILE_FALLBACK)
    cfg = read_db_config(cfg_path)

    with connect(cfg) as conn:
        df_q = fetch_questions(conn, evaluacion_id)
        app_id = _obtener_aplicacion_id(conn, evaluacion_id, informante_id)

        df_excel = pd.read_excel(xls, sheet_name=0)
        df_excel.columns = [str(c).strip() for c in df_excel.columns]

        # Reporte rápido
        stats = print_excel_vs_db_report(df_excel, df_q, col_preg_codigo="PREGUNTA_CODIGO")

        # Mapeo código -> id_pregunta
        map_id = df_q.assign(PREGUNTA_CODIGO=df_q["PREGUNTA_CODIGO"].astype(str).str.strip()).set_index("PREGUNTA_CODIGO")["ID_PREGUNTA"].to_dict()
        df_excel["PreguntaId"] = df_excel["PREGUNTA_CODIGO"].astype(str).str.strip().map(map_id)

        bad = df_excel[df_excel["PreguntaId"].isna()]
        if not bad.empty:
            raise RuntimeError(f"Hay {len(bad)} filas con PREGUNTA_CODIGO sin match. Corrige Excel antes de cargar.")

        # Campos destino
        # NOTA: columnas del Excel provienen del prompt. Ajusta si renombraste encabezados.
        df_excel["SituacionActual"] = df_excel.get("CONVERSACION_INFORMANTE_1","").fillna("").astype(str)
        df_excel["Notas"] = df_excel.get("CONVERSACION_ENTREVISTADOR_1","").fillna("").astype(str)

        # MERGE (insert/update)
        cur = conn.cursor()
        cur.fast_executemany = False

        cur.execute("""
        IF OBJECT_ID('tempdb..#SRC_RESP') IS NOT NULL DROP TABLE #SRC_RESP;
        CREATE TABLE #SRC_RESP(
            PreguntaId INT NOT NULL,
            SituacionActual NVARCHAR(MAX) NULL,
            Notas NVARCHAR(MAX) NULL
        );
        """)

        ins = "INSERT INTO #SRC_RESP(PreguntaId, SituacionActual, Notas) VALUES (?, ?, ?);"
        params = [(int(r.PreguntaId), str(r.SituacionActual), str(r.Notas)) for r in df_excel.itertuples(index=False)]
        cur.executemany(ins, params)

        merge_sql = r"""
        DECLARE @AplicacionId INT = ?;
        MERGE dbo.Respuesta AS T
        USING #SRC_RESP AS S
          ON T.AplicacionId = @AplicacionId
         AND T.PreguntaId   = S.PreguntaId
        WHEN MATCHED THEN
          UPDATE SET
            SituacionActual = CASE
              WHEN S.SituacionActual IS NULL OR LTRIM(RTRIM(S.SituacionActual)) = '' THEN T.SituacionActual
              WHEN T.SituacionActual IS NULL OR LTRIM(RTRIM(T.SituacionActual)) = '' THEN S.SituacionActual
              WHEN CHARINDEX(S.SituacionActual, T.SituacionActual) > 0 THEN T.SituacionActual
              ELSE CONCAT(T.SituacionActual, CHAR(13)+CHAR(10)+CHAR(13)+CHAR(10), S.SituacionActual)
            END,
            Notas = CASE
              WHEN S.Notas IS NULL OR LTRIM(RTRIM(S.Notas)) = '' THEN T.Notas
              WHEN T.Notas IS NULL OR LTRIM(RTRIM(T.Notas)) = '' THEN S.Notas
              WHEN CHARINDEX(S.Notas, T.Notas) > 0 THEN T.Notas
              ELSE CONCAT(T.Notas, CHAR(13)+CHAR(10)+CHAR(13)+CHAR(10), S.Notas)
            END
        WHEN NOT MATCHED BY TARGET THEN
          INSERT (AplicacionId, PreguntaId, SituacionActual, Notas, FechaRegistro)
          VALUES (@AplicacionId, S.PreguntaId, S.SituacionActual, S.Notas, GETDATE());
        """
        cur.execute(merge_sql, (int(app_id),))
        conn.commit()

        stats.update({"aplicacion_id": app_id, "excel_path": str(xls)})
        return stats
