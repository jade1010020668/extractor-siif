"""Uso por línea de comandos (sin interfaz web):

    python -m acta_privada transcripcion.docx -o acta.docx --numero 074 [--sin-ia]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .docx_builder import build_docx
from .llm import DEFAULT_HOST, DEFAULT_MODEL, LLMUnavailable, OllamaClient
from .roster import load_json
from .structure import build_acta
from .transcript import read_any


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="acta_privada")
    ap.add_argument("transcripcion")
    ap.add_argument("-o", "--salida", default="acta.docx")
    ap.add_argument("--numero", default="")
    ap.add_argument("--sin-ia", action="store_true", help="solo reglas, sin Ollama")
    ap.add_argument("--modelo", default=DEFAULT_MODEL)
    ap.add_argument("--host", default=DEFAULT_HOST)
    a = ap.parse_args(argv)

    tr = read_any(a.transcripcion, Path(a.transcripcion).read_bytes())
    llm = None
    if not a.sin_ia:
        llm = OllamaClient(host=a.host, model=a.modelo)
        try:
            if not llm.has_model():
                print(f"El modelo {a.modelo} no está instalado: ollama pull {a.modelo}", file=sys.stderr)
                return 2
        except LLMUnavailable as exc:
            print(exc, file=sys.stderr)
            return 2
    acta = build_acta(tr, load_json("roster"), numero=a.numero, llm=llm,
                      progress=lambda f, m: print(f"[{f:4.0%}] {m}", file=sys.stderr))
    Path(a.salida).write_bytes(build_docx(acta))
    for av in acta.avisos:
        print("AVISO:", av, file=sys.stderr)
    print(f"Acta escrita en {a.salida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
