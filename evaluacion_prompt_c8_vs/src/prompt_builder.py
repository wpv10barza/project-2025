\
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from pathlib import Path
import json
import re
import io

from PyPDF2 import PdfReader


# =========================
# Utilidades PDF / TXT
# =========================

def read_txt(path: Path) -> str:
    if not path.exists():
        raise FileNotFoundError(f"No existe el archivo TXT: {path}")
    return path.read_text(encoding="utf-8", errors="replace").strip()

def read_pdf_text(path: Path) -> str:
    if not path.exists():
        raise FileNotFoundError(f"No existe el PDF: {path}")
    with path.open("rb") as f:
        reader = PdfReader(f)
        chunks = []
        for page in reader.pages:
            chunks.append(page.extract_text() or "")
    return "\n".join(chunks)


# =========================
# Guía: extracción de preguntas
# (robusta pero simple)
# =========================

@dataclass
class Question:
    codigo: str
    descripcion: str
    situacion_deseada: str

def extract_questions_from_guia(guia_pdf: Path) -> List[Question]:
    """
    Intenta extraer (codigo, descripcion, situacion_deseada) desde PREGUNTA_CODIGO.pdf.

    Nota: Los PDFs varían. Esta función busca patrones comunes:
    - Código: 1.1.1 (o 1.2.3)
    - Descripción: ...
    - Situación Deseada: ...
    """
    text = read_pdf_text(guia_pdf)

    # Normaliza saltos de línea
    t = re.sub(r"\r\n?", "\n", text)

    # Patrón flexible: captura bloques por "Código:".
    # Si tu guía usa otro formato, ajusta esta regex.
    block_pat = re.compile(
        r"(?:Código|CODIGO|PREGUNTA_CODIGO)\s*[:\-]?\s*(\d+\.\d+\.\d+)\s*"
        r"(.*?)(?=(?:Código|CODIGO|PREGUNTA_CODIGO)\s*[:\-]?\s*\d+\.\d+\.\d+|\Z)",
        re.DOTALL
    )

    questions: List[Question] = []
    for m in block_pat.finditer(t):
        codigo = m.group(1).strip()
        block = m.group(2).strip()

        # Busca descripción y situación deseada dentro del bloque
        desc = ""
        sit = ""

        desc_m = re.search(r"(?:Descripción|DESCRIPCION|PREGUNTA_DESCRIPCION)\s*[:\-]?\s*(.*?)(?=(?:Situación Deseada|SITUACION DESEADA|PREGUNTA_SITUACION_DESEADA)\s*[:\-]|\Z)", block, re.DOTALL)
        if desc_m:
            desc = desc_m.group(1).strip()

        sit_m = re.search(r"(?:Situación Deseada|SITUACION DESEADA|PREGUNTA_SITUACION_DESEADA)\s*[:\-]?\s*(.*)\Z", block, re.DOTALL)
        if sit_m:
            sit = sit_m.group(1).strip()

        # Fallback: si no hay etiquetas, toma el inicio como descripción
        if not desc:
            # Primeras ~3 líneas como "descripción tentativa"
            lines = [ln.strip() for ln in block.splitlines() if ln.strip()]
            desc = " ".join(lines[:3]).strip()

        questions.append(Question(codigo=codigo, descripcion=desc, situacion_deseada=sit))

    # Dedup por código manteniendo el primero
    seen = set()
    out: List[Question] = []
    for q in questions:
        if q.codigo not in seen:
            out.append(q)
            seen.add(q.codigo)
    return out


# =========================
# Transcripción: extracción L### SPEAKER TEXTO
# =========================

@dataclass
class Turn:
    linea: str
    speaker: str
    texto: str

def extract_turns_from_transcripcion(transcripcion_pdf: Path) -> List[Turn]:
    """
    Extrae turnos con formato tipo:
    L001  E1  Texto...
    """
    text = read_pdf_text(transcripcion_pdf)
    t = re.sub(r"\r\n?", "\n", text)

    # Captura L###, speaker (E1/I1/E2...), y texto hasta el próximo L###
    pat = re.compile(r"L(\d{1,4})\s+(E\d+|I\d+)\s+(.*?)(?=\nL\d{1,4}\s+|\Z)", re.DOTALL)
    turns: List[Turn] = []
    for m in pat.finditer(t):
        linea = m.group(1).zfill(3)
        sp = m.group(2).strip()
        tx = m.group(3).strip()
        # Limpieza mínima
        tx = re.sub(r"\s+\n", "\n", tx)
        tx = re.sub(r"[ \t]{2,}", " ", tx)
        turns.append(Turn(linea=linea, speaker=sp, texto=tx))
    return turns


# =========================
# Roles: carga/creación
# =========================

@dataclass
class RolesConfig:
    # Ejemplo:
    # {
    #  "speakers": {"E1": "Entrevistador principal", "E2": "Entrevistador Marcel", "I1": "Informante 1"},
    #  "entrevistadores": ["E1","E2"],
    #  "informantes": ["I1"]
    # }
    speakers: Dict[str, str]
    entrevistadores: List[str]
    informantes: List[str]

def load_roles(path: Path) -> RolesConfig:
    if not path.exists():
        raise FileNotFoundError(f"No existe roles.json: {path}. Ejecuta: python main.py --init-roles")
    data = json.loads(path.read_text(encoding="utf-8"))
    return RolesConfig(
        speakers=data.get("speakers", {}),
        entrevistadores=data.get("entrevistadores", []),
        informantes=data.get("informantes", []),
    )

def init_roles_from_transcripcion(path_roles: Path, turns: List[Turn]) -> None:
    speakers = sorted({t.speaker for t in turns})
    template = {
        "speakers": {sp: "" for sp in speakers},
        "entrevistadores": [],
        "informantes": [],
        "nota": "Completa speakers con el rol legible. Luego llena entrevistadores/informantes con los IDs (E1/E2/I1...)."
    }
    path_roles.write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")


# =========================
# C8 (FIX): armar PROMPT
# =========================

def build_prompt_c8(
    *,
    objetivo_final_text: str,
    question: Question,
    turns: List[Turn],
    roles: RolesConfig,
) -> str:
    """
    Construye el PROMPT COMPLETO (sin resumir) para revisión.
    Incluye:
    - OBJETIVO_FINAL (texto completo)
    - Datos de la pregunta (código/descr/situación)
    - Columnas/roles
    - Transcripción estructurada: LÍNEA\tSPEAKER\tROL\tTEXTO
    """
    # Columnas dinámicas (según roles.json)
    entrevistadores = roles.entrevistadores[:] or [sp for sp in roles.speakers.keys() if sp.startswith("E")]
    informantes = roles.informantes[:] or [sp for sp in roles.speakers.keys() if sp.startswith("I")]

    col_entrev = [f"CONVERSACION_ENTREVISTADOR_{i+1}" for i in range(len(entrevistadores))]
    col_info = [f"CONVERSACION_INFORMANTE_{i+1}" for i in range(len(informantes))]

    columnas = ["PREGUNTA_CODIGO"] + col_entrev + col_info + ["PREGUNTA_DESCRIPCION", "PREGUNTA_SITUACION_DESEADA", "ANALISIS_DE_COBERTURA"]
    header = "\t".join(columnas)

    # Mapeo speaker -> columna (solo para explicar asignación)
    asignacion_lines = []
    for i, sp in enumerate(entrevistadores):
        asignacion_lines.append(f"- {col_entrev[i]}: usa SOLO intervenciones de {sp} ({roles.speakers.get(sp,'ROL_NO_DEFINIDO')}).")
    for i, sp in enumerate(informantes):
        asignacion_lines.append(f"- {col_info[i]}: usa SOLO intervenciones de {sp} ({roles.speakers.get(sp,'ROL_NO_DEFINIDO')}).")

    # Transcripción estructurada (tabla)
    table_lines = []
    table_lines.append("Formato: LÍNEA\tSPEAKER\tROL\tTEXTO")
    table_lines.append("")
    for tr in turns:
        rol = roles.speakers.get(tr.speaker, "ROL_NO_DEFINIDO")
        # No modificar texto: solo normalize saltos en una línea para tabla, manteniendo literal lo más posible
        texto = tr.texto.replace("\t", " ").strip()
        table_lines.append(f"L{tr.linea}\t{tr.speaker}\t{rol}\t{texto}")

    transcripcion_struct = "\n".join(table_lines)

    prompt = f"""\
--- OBJETIVO_FINAL (texto completo) ---
{objetivo_final_text}

AHORA TRABAJA SOLO CON ESTE CÓDIGO:
PREGUNTA_CODIGO: {question.codigo}
PREGUNTA_DESCRIPCION: {question.descripcion}
PREGUNTA_SITUACION_DESEADA: {question.situacion_deseada}

COLUMNAS Y ROLES (EN ORDEN OBLIGATORIO):
{header}

Asignación de speakers a columnas:
{chr(10).join(asignacion_lines)}

FUENTE ÚNICA DE EVIDENCIA TEXTUAL (TABLA):
------------------ INICIO TRANSCRIPCION ESTRUCTURADA ------------------
{transcripcion_struct}
------------------ FIN TRANSCRIPCION ESTRUCTURADA ------------------

INSTRUCCIONES:
- Devuelve SOLO UNA LÍNEA TSV (sin encabezado) siguiendo exactamente el orden de columnas de:
  {header}
- Cada celda de conversación debe contener SOLO citas textuales (puedes concatenar varios fragmentos).
- No inventes nada. Si no está cubierta, deja celdas vacías y marca NO CUBIERTA TEXTUALMENTE con justificación.
- No añadas texto extra antes ni después de la línea TSV.
"""
    return prompt.strip() + "\n"
