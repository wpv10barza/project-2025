\
from __future__ import annotations

import argparse
from pathlib import Path
import sys

from src.config import (
    OBJETIVO_FINAL_TXT,
    GUIA_PREGUNTA_CODIGO_PDF,
    TRANSCRIPCION_PDF,
    ROLES_JSON,
    OUT_DIR,
)

from src.prompt_builder import (
    extract_questions_from_guia,
    extract_turns_from_transcripcion,
    init_roles_from_transcripcion,
    load_roles,
    build_prompt_c8,
    Question,
)

def ensure_out_dir() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

def save_prompt(codigo: str, prompt: str) -> Path:
    ensure_out_dir()
    out_path = OUT_DIR / f"prompt_C8_{codigo}.txt"
    out_path.write_text(prompt, encoding="utf-8")
    return out_path

def main() -> int:
    parser = argparse.ArgumentParser(description="C8 (FIX) - Genera PROMPT completo (guía + objetivo + transcripción) y lo guarda en TXT.")
    parser.add_argument("--codigo", type=str, help="Código específico (ej: 1.1.1)")
    parser.add_argument("--all", action="store_true", help="Generar prompt para todos los códigos detectados")
    parser.add_argument("--init-roles", action="store_true", help="Crear roles.json basado en speakers detectados en la transcripción")
    args = parser.parse_args()

    # 1) Leer transcripción y extraer speakers
    turns = extract_turns_from_transcripcion(TRANSCRIPCION_PDF)
    if not turns:
        print("❌ No se detectaron turnos en Transcripcion.pdf. Revisa que el PDF contenga líneas tipo 'L001 E1 ...'.")
        return 2

    # 2) Init roles si se pidió
    if args.init_roles:
        init_roles_from_transcripcion(ROLES_JSON, turns)
        print(f"✅ roles.json creado en: {ROLES_JSON}")
        print("➡️ Edita roles.json (roles, entrevistadores, informantes) y vuelve a ejecutar.")
        return 0

    # 3) Cargar roles
    roles = load_roles(ROLES_JSON)

    # 4) Leer objetivo final
    objetivo_final_text = OBJETIVO_FINAL_TXT.read_text(encoding="utf-8", errors="replace").strip()
    if not objetivo_final_text:
        print("❌ OBJETIVO_FINAL.txt está vacío.")
        return 2

    # 5) Leer guía y extraer preguntas
    questions = extract_questions_from_guia(GUIA_PREGUNTA_CODIGO_PDF)
    if not questions:
        print("❌ No se detectaron preguntas en PREGUNTA_CODIGO.pdf con el parser actual.")
        print("   Puedes ajustar la regex en src/prompt_builder.py -> extract_questions_from_guia().")
        return 2

    q_by_code = {q.codigo: q for q in questions}

    # 6) Generación
    if args.codigo:
        code = args.codigo.strip()
        if code not in q_by_code:
            print(f"❌ Código no encontrado en guía: {code}")
            print(f"   Códigos detectados: {', '.join(list(q_by_code.keys())[:25])}{'...' if len(q_by_code)>25 else ''}")
            return 2

        q = q_by_code[code]
        prompt = build_prompt_c8(objetivo_final_text=objetivo_final_text, question=q, turns=turns, roles=roles)
        out = save_prompt(code, prompt)
        print(f"✅ Prompt C8 guardado en: {out}")
        return 0

    if args.all:
        ok = 0
        for q in questions:
            prompt = build_prompt_c8(objetivo_final_text=objetivo_final_text, question=q, turns=turns, roles=roles)
            out = save_prompt(q.codigo, prompt)
            ok += 1
        print(f"✅ Generados {ok} prompts en: {OUT_DIR}")
        return 0

    print("Uso:")
    print("  python main.py --init-roles")
    print("  python main.py --codigo 1.1.1")
    print("  python main.py --all")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
