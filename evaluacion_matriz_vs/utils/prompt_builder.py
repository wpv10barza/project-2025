from __future__ import annotations

from typing import Dict, List


def build_prompt(
    *,
    question_code: str,
    question_desc: str,
    question_desired: str,
    objective_final_text: str,
    structured_transcript_tsv: str,
    interviewer_speakers: List[str],
    informant_speakers: List[str],
    speaker_roles: Dict[str, str],
) -> str:
    """Construye el prompt para una sola pregunta.

    Importante:
    - objective_final_text se incluye literal.
    - structured_transcript_tsv es la única evidencia.
    """

    # Columnas dinámicas:
    interviewer_cols = [f"CONVERSACION_ENTREVISTADOR_{i+1}" for i in range(len(interviewer_speakers))]
    informant_cols = [f"CONVERSACION_INFORMANTE_{i+1}" for i in range(len(informant_speakers))]
    columns = ["PREGUNTA_CODIGO", *interviewer_cols, *informant_cols, "PREGUNTA_DESCRIPCION", "PREGUNTA_SITUACION_DESEADA", "ANALISIS_DE_COBERTURA"]

    # Mapeo speaker -> columna
    assign_lines = []
    for i, spk in enumerate(interviewer_speakers, start=1):
        assign_lines.append(f"- CONVERSACION_ENTREVISTADOR_{i}: usa SOLO intervenciones de {spk} ({speaker_roles.get(spk,'')}).")
    for i, spk in enumerate(informant_speakers, start=1):
        assign_lines.append(f"- CONVERSACION_INFORMANTE_{i}: usa SOLO intervenciones de {spk} ({speaker_roles.get(spk,'')}).")

    prompt = f"""{objective_final_text}

AHORA TRABAJA SOLO CON ESTE CÓDIGO:
PREGUNTA_CODIGO: {question_code}
PREGUNTA_DESCRIPCION: {question_desc}
PREGUNTA_SITUACION_DESEADA: {question_desired}

COLUMNAS Y ROLES (EN ORDEN OBLIGATORIO):
{'\t'.join(columns)}

Asignación de speakers a columnas:
{chr(10).join(assign_lines)}

FUENTE ÚNICA DE EVIDENCIA TEXTUAL (TABLA):
Formato: LÍNEA\tSPEAKER\tROL\tTEXTO

------------------ INICIO TRANSCRIPCION ESTRUCTURADA ------------------
{structured_transcript_tsv}
------------------ FIN TRANSCRIPCION ESTRUCTURADA ------------------

INSTRUCCIONES:
- Devuelve SOLO UNA LÍNEA TSV (sin encabezado) siguiendo exactamente el orden de columnas:
  {'\t'.join(columns)}
- Cada celda de conversación debe contener SOLO citas textuales (puedes concatenar varios fragmentos; si quieres separarlos usa ' | ' o saltos de línea).
- NO inventes nada. Si no está cubierta, deja celdas de conversación vacías.
- En ANALISIS_DE_COBERTURA escribe exactamente una de:
  TOTALMENTE CUBIERTA | PARCIALMENTE CUBIERTA | NO CUBIERTA TEXTUALMENTE
  y agrega una justificación MUY CORTA basada únicamente en citas textuales que YA estén copiadas en las celdas de conversación (no agregues evidencia nueva ahí).
- No añadas texto extra antes ni después de la línea TSV.
"""
    return prompt
