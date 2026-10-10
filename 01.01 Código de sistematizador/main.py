import sys
from src.prompt_builder import write_prompt_files
from src.gpt_runner import run_gpt_and_save
from config import RESULT_RAW_TXT

def main() -> None:
    cmd = (sys.argv[1].lower() if len(sys.argv) > 1 else "preview").strip()

    if cmd in ("preview", "c8", "prompt"):
        out = write_prompt_files()
        print(f"[OK] Prompt generado para revisar: {out}")
        return

    if cmd in ("run", "gpt", "execute"):
        run_gpt_and_save()
        print(f"[OK] Resultado RAW: {RESULT_RAW_TXT}")
        print("[OK] Además se intentó guardar TSV/CSV/XLSX en 01.03 Resultado ChatGPT")
        return

    print("Uso:")
    print("  python main.py preview   # genera prompt_completo.txt (C8 completo) en 01.02 Promp")
    print("  python main.py run       # llama a GPT-4o y guarda resultados en 01.03 Resultado ChatGPT")
    sys.exit(2)

if __name__ == "__main__":
    main()
