"""Ubicación de los datos del usuario (fuera de la carpeta del programa).

Windows: %LOCALAPPDATA%\\ActaPrivada   ·   otros: ~/.acta-privada
Se puede cambiar con la variable ACTA_DATA_DIR.
"""
from __future__ import annotations

import json
import os
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent.parent


def data_dir() -> Path:
    env = os.environ.get("ACTA_DATA_DIR")
    if env:
        base = Path(env)
    elif os.name == "nt" and os.environ.get("LOCALAPPDATA"):
        base = Path(os.environ["LOCALAPPDATA"]) / "ActaPrivada"
    else:
        base = Path.home() / ".acta-privada"
    return base


def user_config(name: str) -> Path:
    return data_dir() / "config" / f"{name}.json"


def user_plantilla() -> Path:
    return data_dir() / "plantilla" / "plantilla_acta.docx"


def modelos_dir() -> Path:
    return data_dir() / "modelos"


def leer_ajustes() -> dict:
    try:
        return json.loads((data_dir() / "ajustes.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def guardar_ajustes(**cambios) -> None:
    aj = leer_ajustes() | cambios
    escribir_json(data_dir() / "ajustes.json", aj)


def escribir_json(path: Path, obj) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)
