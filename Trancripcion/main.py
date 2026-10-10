from __future__ import annotations
import argparse
from pathlib import Path
import os

from src.utils import ensure_dirs, print_err
from src.whisper_transcriber import transcribe_missing_mp3
from src.openai_cleaner import clean_txt_file
from src.prompt_builder import build_prompt_from_clean

def cmd_transcribe(args) -> int:
    dirs = ensure_dirs()
    created = transcribe_missing_mp3(dirs["audio"], dirs["texto"], model_name=args.model)
    print(f"[OK] Whisper creó {len(created)} TXT(s) en {dirs['texto']}")
    return 0

def cmd_clean(args) -> int:
    dirs = ensure_dirs()
    model = args.model or os.getenv("OPENAI_MODEL", "gpt-5.2")

    txts = sorted(dirs["texto"].glob("*.txt"))
    if not txts:
        print("[OpenAI] No hay TXT en /texto. Ejecuta primero: python main.py transcribe")
        return 0

    n_ok = 0
    for t in txts:
        out = dirs["limpios"] / f"{t.stem}_LIMPIO.txt"
        if out.exists() and not args.force:
            print(f"[OpenAI] Skip (ya existe): {out.name}")
            continue
        r = clean_txt_file(t, out, model=model, max_retries=args.retries)
        if r:
            n_ok += 1

    print(f"[OK] OpenAI generó {n_ok} TXT(s) en {dirs['limpios']}")
    return 0

def cmd_prompt(args) -> int:
    dirs = ensure_dirs()
    clean_files = sorted(dirs["limpios"].glob("*_LIMPIO.txt"))
    if not clean_files:
        print("[Prompt] No hay archivos limpios en /limpios. Ejecuta: python main.py clean")
        return 0

    n = 0
    for c in clean_files:
        out = dirs["prompts"] / f"PROMPT_FINAL_{c.stem.replace('_LIMPIO','')}.txt"
        if out.exists() and not args.force:
            print(f"[Prompt] Skip (ya existe): {out.name}")
            continue
        build_prompt_from_clean(c, out)
        print(f"[Prompt] OK -> {out.name}")
        n += 1

    print(f"[OK] Prompts generados: {n} en {dirs['prompts']}")
    return 0

def cmd_all(args) -> int:
    # 1) whisper
    class A: pass
    a = A(); a.model = args.whisper_model
    cmd_transcribe(a)

    # 2) openai
    b = A(); b.model = args.openai_model; b.force = args.force; b.retries = args.retries
    cmd_clean(b)

    # 3) prompt
    c = A(); c.force = args.force
    cmd_prompt(c)
    return 0

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="ai-transcripcion",
        description="MP3 -> Whisper local -> Limpieza OpenAI -> Prompt final",
    )
    sub = p.add_subparsers(dest="cmd", required=False)

    p_tr = sub.add_parser("transcribe", help="Transcribe MP3s en /audio hacia /texto")
    p_tr.add_argument("--model", default="medium", help="Whisper model (tiny/base/small/medium/large)")
    p_tr.set_defaults(func=cmd_transcribe)

    p_cl = sub.add_parser("clean", help="Estructura TXT con OpenAI hacia /limpios")
    p_cl.add_argument("--model", default=None, help="OpenAI model (default: OPENAI_MODEL o gpt-5.2)")
    p_cl.add_argument("--force", action="store_true", help="Reprocesa aunque exista _LIMPIO")
    p_cl.add_argument("--retries", type=int, default=3, help="Reintentos ante error (429, etc.)")
    p_cl.set_defaults(func=cmd_clean)

    p_pr = sub.add_parser("prompt", help="Crea prompt final desde /limpios hacia /prompts")
    p_pr.add_argument("--force", action="store_true", help="Reescribe prompt aunque exista")
    p_pr.set_defaults(func=cmd_prompt)

    p_all = sub.add_parser("all", help="Ejecuta transcribe + clean + prompt")
    p_all.add_argument("--whisper-model", default="medium")
    p_all.add_argument("--openai-model", default=None)
    p_all.add_argument("--force", action="store_true")
    p_all.add_argument("--retries", type=int, default=3)
    p_all.set_defaults(func=cmd_all)

    return p

def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    # default: si no pasan comando, muestra ayuda
    if not getattr(args, "cmd", None):
        parser.print_help()
        print("\nEjemplo:\n  python main.py transcribe\n  python main.py clean\n  python main.py prompt\n")
        return 0

    try:
        return args.func(args)
    except Exception as e:
        print_err(e)
        return 1

if __name__ == "__main__":
    raise SystemExit(main())
