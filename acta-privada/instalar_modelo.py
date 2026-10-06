"""Descarga el modelo de IA durante la instalación (lo ejecuta el instalador).

Uso:  python instalar_modelo.py --perfil auto|8gb|16gb   [--modelo NOMBRE]

* Inicia el Ollama incluido (solo 127.0.0.1) si no hay uno abierto.
* Elige el modelo según la RAM del equipo (perfil «auto») y lo descarga.
* Solo se bajan archivos del catálogo de modelos; no se envía ningún contenido.
* Si no hay internet NO falla la instalación: la app permite descargarlo después.
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from acta_privada.llm import DEFAULT_HOST, LLMUnavailable, OllamaClient  # noqa: E402
from acta_privada.paths import data_dir  # noqa: E402
from acta_privada.perfiles import PERFILES, ram_gb, recomendado  # noqa: E402


def descargar(modelo: str, host: str = DEFAULT_HOST, log=print) -> bool:
    """True si el modelo queda disponible."""
    client = OllamaClient(host=host, model=modelo)
    try:
        if client.has_model():
            log(f"El modelo {modelo} ya está instalado.")
            return True
        ultimo = {"p": -10}

        def progreso(frac, estado):
            if frac is None:
                if estado and estado != ultimo.get("e"):
                    log(f"  {estado}")
                    ultimo["e"] = estado
                return
            pct = int(frac * 100)
            if pct >= ultimo["p"] + 5 or pct == 100:
                ultimo["p"] = pct
                log(f"  Descargando {modelo}: {pct}%")

        client.pull(progress=progreso)
        log(f"Modelo {modelo} instalado.")
        return True
    except LLMUnavailable as exc:
        log(f"\nNo se pudo descargar el modelo: {exc}")
        log("No es grave: abra Actas Privadas y pulse «Descargar modelo» cuando tenga internet.")
        return False


def main(argv=None) -> int:
    for s in (sys.stdout, sys.stderr):
        if hasattr(s, "reconfigure"):
            s.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("--perfil", choices=["auto", *sorted(PERFILES)], default="auto")
    ap.add_argument("--modelo", help="anula el modelo del perfil (pruebas)")
    ap.add_argument("--host", default=DEFAULT_HOST)
    ap.add_argument("--sin-pausa", action="store_true")
    a = ap.parse_args(argv)

    ram = ram_gb()
    clave = recomendado(ram) if a.perfil == "auto" else a.perfil
    perfil = PERFILES[clave]
    modelo = a.modelo or perfil.modelo
    print("=" * 62)
    print("  ACTAS PRIVADAS — instalando el modelo de IA (una sola vez)")
    if ram:
        print(f"  Memoria del equipo: {ram:.0f} GB")
    print(f"  Modelo: {modelo} (≈{perfil.descarga_gb:.1f} GB). No cierre esta ventana.")
    print("=" * 62)

    import launcher                                   # inicia el Ollama incluido
    data_dir().mkdir(parents=True, exist_ok=True)
    ollama = launcher.iniciar_ollama()
    ok = False
    try:
        ok = descargar(modelo, a.host)
        if ok:
            from acta_privada.paths import guardar_ajustes
            guardar_ajustes(perfil=clave)
    finally:
        if ollama and ollama.poll() is None:
            ollama.terminate()
            try:
                ollama.wait(10)
            except Exception:                         # noqa: BLE001
                ollama.kill()
    if not ok and not a.sin_pausa:
        time.sleep(12)                                # que el usuario alcance a leer
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
