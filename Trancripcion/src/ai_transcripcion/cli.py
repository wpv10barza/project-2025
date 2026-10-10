from __future__ import annotations
import argparse
from pathlib import Path
from openai import OpenAI

def cmd_transcribe(args: argparse.Namespace) -> int:
    inp = Path(args.input).expanduser().resolve()
    out = Path(args.output).expanduser().resolve()

    if not inp.exists():
        raise FileNotFoundError(f"No existe el archivo: {inp}")

    out.parent.mkdir(parents=True, exist_ok=True)

    client = OpenAI()
    with inp.open("rb") as f:
        r = client.audio.transcriptions.create(
            model=args.model,
            file=f,
            response_format="text" if args.format == "text" else "json",
        )

    text = r if isinstance(r, str) else getattr(r, "text", None) or str(r)
    out.write_text(text, encoding="utf-8")
    print(f"OK -> {out}")
    return 0

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="ai-transcripcion")
    sub = p.add_subparsers(dest="cmd", required=True)

    t = sub.add_parser("transcribe", help="Transcribe audio (mp3/wav/m4a/etc) a texto")
    t.add_argument("-i", "--input", required=True, help="Ruta del audio (mp3/wav/m4a/...)")
    t.add_argument("-o", "--output", required=True, help="Ruta del TXT de salida")
    t.add_argument("--model", default="gpt-4o-transcribe",
                   help="gpt-4o-transcribe | gpt-4o-mini-transcribe | whisper-1")
    t.add_argument("--format", choices=["text", "json"], default="text")
    t.set_defaults(func=cmd_transcribe)
    return p

def main() -> int:
    p = build_parser()
    args = p.parse_args()
    return args.func(args)

if __name__ == "__main__":
    raise SystemExit(main())
