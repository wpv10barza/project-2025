from __future__ import annotations

import re
from typing import Dict, List, Optional


# Código tipo 1.1.1 / 1.2.3.4 etc.
CODE_PATTERN = r"(?P<code>\d+(?:\.\d+)+)"

# Intento principal: PDFs con etiquetas "Código:", "Descripción:", "Situación Deseada:"
LABELED_BLOCK_RE = re.compile(
    rf"""(?isx)
    C[oó]digo\s*[:\-]\s*{CODE_PATTERN}\s*
    (?:\n|\r|\s)+
    Descripci[oó]n\s*[:\-]\s*(?P<desc>.*?)
    (?:
        (?:\n|\r|\s)+
        Situaci[oó]n\s*Deseada\s*[:\-]\s*(?P<desired>.*?)
    )?
    (?=\n\s*C[oó]digo\s*[:\-]\s*\d+(?:\.\d+)+|\Z)
    """
)

# Fallback: bloques iniciados por un código al comienzo de línea (cuando no existe "Código:")
FALLBACK_CODE_LINE_RE = re.compile(rf"(?m)^\s*{CODE_PATTERN}\s+")


def _clean_cell_text(s: str) -> str:
    # Limpieza mínima: no parafrasea, solo normaliza espacios extremos
    s = s.strip()
    # Evita saltos de línea múltiples que rompan demasiado la celda
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s


def parse_questions(text: str) -> List[Dict[str, str]]:
    """Extrae preguntas con campos:
    - PREGUNTA_CODIGO
    - PREGUNTA_DESCRIPCION
    - PREGUNTA_SITUACION_DESEADA

    Si no puede detectar formato etiquetado, usa un fallback por código al inicio de línea.
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    out: List[Dict[str, str]] = []

    labeled = list(LABELED_BLOCK_RE.finditer(text))
    if labeled:
        for m in labeled:
            code = m.group("code").strip()
            desc = _clean_cell_text(m.group("desc") or "")
            desired = _clean_cell_text(m.group("desired") or "")
            out.append({
                "PREGUNTA_CODIGO": code,
                "PREGUNTA_DESCRIPCION": desc,
                "PREGUNTA_SITUACION_DESEADA": desired,
            })
        return out

    # ===== Fallback =====
    starts = list(FALLBACK_CODE_LINE_RE.finditer(text))
    if not starts:
        return out

    # Cortamos por posiciones
    for i, m in enumerate(starts):
        code = m.group("code").strip()
        start = m.start()
        end = starts[i + 1].start() if i + 1 < len(starts) else len(text)
        block = text[start:end].strip()

        # Heurística: primera línea después del código como descripción
        # Buscamos si hay etiquetas dentro
        desc = ""
        desired = ""

        # Si dentro del bloque existen etiquetas, úsalas:
        dm = re.search(r"(?is)Descripci[oó]n\s*[:\-]\s*(.*?)(?=Situaci[oó]n\s*Deseada\s*[:\-]|\Z)", block)
        if dm:
            desc = _clean_cell_text(dm.group(1))
        else:
            # Si no hay etiqueta, usa el bloque completo sin la primera línea del código
            desc = _clean_cell_text(re.sub(rf"^\s*{re.escape(code)}\s+", "", block))

        sm = re.search(r"(?is)Situaci[oó]n\s*Deseada\s*[:\-]\s*(.*)\Z", block)
        if sm:
            desired = _clean_cell_text(sm.group(1))

        out.append({
            "PREGUNTA_CODIGO": code,
            "PREGUNTA_DESCRIPCION": desc,
            "PREGUNTA_SITUACION_DESEADA": desired,
        })

    return out


def merge_question_lists(lists: List[List[Dict[str, str]]]) -> List[Dict[str, str]]:
    """Une múltiples listas de preguntas por PREGUNTA_CODIGO.
    Regla: conserva el texto más largo no vacío para descripción/situación.
    """
    by_code: Dict[str, Dict[str, str]] = {}
    for lst in lists:
        for q in lst:
            code = q.get("PREGUNTA_CODIGO", "").strip()
            if not code:
                continue
            if code not in by_code:
                by_code[code] = {
                    "PREGUNTA_CODIGO": code,
                    "PREGUNTA_DESCRIPCION": q.get("PREGUNTA_DESCRIPCION", "").strip(),
                    "PREGUNTA_SITUACION_DESEADA": q.get("PREGUNTA_SITUACION_DESEADA", "").strip(),
                }
            else:
                for k in ["PREGUNTA_DESCRIPCION", "PREGUNTA_SITUACION_DESEADA"]:
                    cur = by_code[code].get(k, "")
                    new = q.get(k, "").strip()
                    if len(new) > len(cur):
                        by_code[code][k] = new

    # Orden por código numérico (1.2.10 > 1.2.2)
    def _key(c: str):
        try:
            return [int(x) for x in c.split('.')]
        except Exception:
            return [999999]

    return [by_code[c] for c in sorted(by_code.keys(), key=_key)]
