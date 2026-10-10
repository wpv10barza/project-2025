SQL_PREGUNTAS = """
SELECT *
FROM dbo.VW_PREGUNTA_POR_EVALUACION
WHERE EvaluacionId = ?
ORDER BY 1;
"""

SQL_APLICACION = """
SELECT TOP (1) Id
FROM dbo.Aplicacion
WHERE EvaluacionId = ? AND InformanteId = ?
ORDER BY Id DESC;
"""

SQL_EXISTING_RESP_COUNT = """
SELECT COUNT(*) AS n
FROM dbo.Respuesta
WHERE AplicacionId = ?;
"""

SQL_EXISTING_PREGUNTAS = """
SELECT PreguntaId
FROM dbo.Respuesta
WHERE AplicacionId = ?;
"""
