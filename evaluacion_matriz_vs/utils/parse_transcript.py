from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable, List, Dict, Tuple


@dataclass
class Turn:
    line_id: str           # e.g., "L001" or "001" -> we keep as "L001"
    speaker: str           # e.g., "E1", "E2", "I1"
    text: str              # literal text (as extracted)
    time_span: str = ""    # optional


# Captura:
# - L### (o L###, L###...)  -> line id(s)
# - speaker (E\d+ o I\d+)
# - texto hasta antes del siguiente L### o fin
_TURN_PATTERN = re.compile(
    r"""(?mxs)
    ^\s*L(?P<line_ids>\d+(?:\s*,\s*L\d+)*)      # L001 or L001, L002
    \s+(?P<speaker>[A-Za-z]+\d+)                  # E1 / I1 / E2 ...
    \s+(?P<body>.*?)
    (?=^\s*L\d+\s+[A-Za-z]+\d+\s+|\Z)
    """
)

_TIME_SPAN_PATTERN = re.compile(r"(\d{2}:\d{2}:\d{2},\d{3}\s*–\s*\d{2}:\d{2}:\d{2},\d{3})")


def parse_turns(transcription_text: str) -> List[Turn]:
    turns: List[Turn] = []
    for m in _TURN_PATTERN.finditer(transcription_text):
        line_ids_raw = m.group("line_ids").strip()
        speaker = m.group("speaker").strip()
        body = (m.group("body") or "").strip()

        time_span = ""
        ts = _TIME_SPAN_PATTERN.search(body)
        if ts:
            time_span = ts.group(1).strip()
            body = body[:ts.start()].strip()

        # Normaliza line_ids_raw a una forma consistente "L###" o "L###,L###"
        # Manteniendo literalidad en el contenido, pero el identificador lo ordenamos.
        # Si viene "001" lo convertimos a "L001".
        # Si viene "001, L002" -> "L001,L002"
        line_ids = []
        for token in re.split(r"\s*,\s*", line_ids_raw):
            token = token.strip()
            token = token[1:] if token.upper().startswith("L") else token
            token = token.zfill(3)
            line_ids.append("L" + token)

        line_id_norm = ",".join(line_ids)
        turns.append(Turn(line_id=line_id_norm, speaker=speaker, text=body, time_span=time_span))

    return turns


def turns_to_structured_tsv(turns: List[Turn], speaker_roles: Dict[str, str]) -> str:
    """Genera una tabla TSV: LINEA\tSPEAKER\tROL\tTEXTO (una fila por Turn)."""
    lines = []
    for t in turns:
        rol = speaker_roles.get(t.speaker, "")
        # NO modificamos t.text; solo garantizamos que no haya tabs que rompan el TSV
        texto = t.text.replace("\t", " ")
        lines.append(f"{t.line_id}\t{t.speaker}\t{rol}\t{texto}")
    return "\n".join(lines)
