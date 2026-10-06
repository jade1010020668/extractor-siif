"""Arranque de la app empaquetada (doble clic en «Actas Privadas»).

1. Inicia el Ollama incluido, escuchando SOLO en 127.0.0.1, con los modelos en
   la carpeta de datos del usuario (si ya hay un Ollama local abierto, lo usa).
2. Inicia la interfaz en 127.0.0.1 y abre el navegador.
3. Al cerrar la ventana negra se detiene todo.

`--autotest` arranca todo, verifica que responda y sale (lo usa la compilación).
"""
from __future__ import annotations

import os
import socket
import subprocess
import sys
import time
import urllib.request
import webbrowser
from pathlib import Path

APP = Path(__file__).resolve().parent
sys.path.insert(0, str(APP))
from acta_privada.paths import data_dir, modelos_dir  # noqa: E402

OLLAMA_PORT = 11434
_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def responde(url: str, timeout: float = 2) -> bool:
    try:
        with _OPENER.open(url, timeout=timeout) as r:
            return r.status == 200
    except Exception:                                     # noqa: BLE001
        return False


def esperar(url: str, segundos: float) -> bool:
    fin = time.time() + segundos
    while time.time() < fin:
        if responde(url):
            return True
        time.sleep(0.5)
    return False


def puerto_libre(desde: int = 8501) -> int:
    for port in range(desde, desde + 50):
        with socket.socket() as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("No hay puertos libres para la interfaz.")


def ollama_incluido() -> Path | None:
    exe = APP / "ollama" / ("ollama.exe" if os.name == "nt" else "ollama")
    return exe if exe.exists() else None


def iniciar_ollama() -> subprocess.Popen | None:
    url = f"http://127.0.0.1:{OLLAMA_PORT}/api/tags"
    if responde(url):
        print("· Usando el Ollama que ya está abierto en este equipo.")
        return None
    exe = ollama_incluido()
    if not exe:
        print("· No se encontró Ollama: la app funcionará en modo «Solo reglas».")
        return None
    modelos_dir().mkdir(parents=True, exist_ok=True)
    env = os.environ | {
        "OLLAMA_HOST": f"127.0.0.1:{OLLAMA_PORT}",   # solo este equipo
        "OLLAMA_MODELS": str(modelos_dir()),
        "OLLAMA_KEEP_ALIVE": "10m",
    }
    proc = subprocess.Popen([str(exe), "serve"], env=env,
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if esperar(url, 60):
        print("· IA local (Ollama) lista.")
    else:
        print("· Ollama no respondió; la app funcionará en modo «Solo reglas».")
    return proc


def main() -> int:
    for stream in (sys.stdout, sys.stderr):           # consolas Windows sin UTF-8
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")
    autotest = "--autotest" in sys.argv
    data_dir().mkdir(parents=True, exist_ok=True)
    print("=" * 62)
    print("  ACTAS PRIVADAS — todo se procesa en este equipo")
    print("  No cierre esta ventana mientras use la aplicación.")
    print("=" * 62)
    ollama = iniciar_ollama()
    port = puerto_libre()
    env = os.environ | {"ACTA_DATA_DIR": str(data_dir()),
                        "STREAMLIT_BROWSER_GATHER_USAGE_STATS": "false"}
    ui = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", str(APP / "app.py"),
         "--server.address", "127.0.0.1", "--server.port", str(port),
         "--server.headless", "true", "--browser.gatherUsageStats", "false"],
        cwd=str(APP), env=env)
    url = f"http://127.0.0.1:{port}"
    try:
        if not esperar(url + "/_stcore/health", 90):
            print("ERROR: la interfaz no arrancó.")
            return 1
        print(f"· Aplicación abierta en {url}")
        if autotest:
            ok_ollama = ollama is None or responde(f"http://127.0.0.1:{OLLAMA_PORT}/api/tags")
            print("AUTOTEST", "OK" if ok_ollama else "FALLO (ollama)")
            return 0 if ok_ollama else 1
        webbrowser.open(url)
        ui.wait()
        return 0
    except KeyboardInterrupt:
        return 0
    finally:
        for p in (ui, ollama):
            if p and p.poll() is None:
                p.terminate()
                try:
                    p.wait(10)
                except subprocess.TimeoutExpired:
                    p.kill()


if __name__ == "__main__":
    raise SystemExit(main())
