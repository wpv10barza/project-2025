import os
import sys
from src.prompt_builder import write_prompt_files
from config import RESULT_RAW_TXT

def main() -> None:
    # ✅ v6: si hay OPENAI_API_KEY, por default corre RUN (y genera C8 dentro).
    # Si NO hay key, por default corre PREVIEW.
    if len(sys.argv) > 1:
        cmd = sys.argv[1].lower().strip()
    else:
        cmd = "run" if os.getenv("OPENAI_API_KEY") else "preview"

    if cmd in ("preview", "c8", "prompt"):
        out = write_prompt_files()
        print(f"[OK] Prompt generado para revisar: {out}")
        print("[TIP] Para ejecutar GPT y generar resultados: setea OPENAI_API_KEY y corre: python main.py run")
        return

    if cmd in ("run", "gpt", "execute"):
        # Import lazy: preview NO requiere openai
        from src.gpt_runner import run_gpt_and_save
        run_gpt_and_save()
        print(f"[OK] Resultado RAW: {RESULT_RAW_TXT}")
        print("[OK] Además se guardó TSV/CSV/XLSX en 01.03 Resultado ChatGPT")
        return

    if cmd in ("doctor", "diag", "diagnostico"):
        from src.diagnostics import run_diagnostics
        ok = run_diagnostics()
        raise SystemExit(0 if ok else 1)

    print("Uso:")
    print("  python main.py preview     # genera prompt_completo.txt (C8) en 01.02 Promp")
    print("  python main.py run         # llama al modelo y guarda resultados en 01.03 Resultado ChatGPT")
    print("  python main.py doctor      # diagnóstico rápido de entorno/insumos")
    raise SystemExit(2)

if __name__ == "__main__":
    main()
