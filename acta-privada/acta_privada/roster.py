"""Padrón de personas (nombres canónicos, cargos). Se guarda en un archivo
local fuera de git; también puede editarse en la interfaz."""
from __future__ import annotations

import difflib
import json
import re
import unicodedata
from pathlib import Path

from .models import Person

CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"


def norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c)).lower()
    s = re.sub(r"\b(dr|dra|doctor|doctora|dr\.|dra\.)\b", " ", s)
    return " ".join(re.sub(r"[^a-z0-9 ]+", " ", s).split())


def load_json(name: str) -> dict:
    """Carga config/<name>.json y, si no existe, config/<name>.example.json."""
    for fname in (f"{name}.json", f"{name}.example.json"):
        p = CONFIG_DIR / fname
        if p.exists():
            return json.loads(p.read_text(encoding="utf-8"))
    return {}


def _score(raw: str, candidate: str) -> float:
    a, b = norm(raw), norm(candidate)
    if not a or not b:
        return 0.0
    ta, tb = set(a.split()), set(b.split())
    if ta <= tb or tb <= ta:                      # "Ana Pérez" ⊂ "Ana María Pérez Gómez"
        return 0.95 if min(len(ta), len(tb)) >= 2 else 0.85
    return difflib.SequenceMatcher(None, a, b).ratio()


def match_person(raw: str, roster: dict, umbral: float = 0.8) -> Person:
    best, best_score = None, 0.0
    for p in roster.get("personas", []):
        for cand in [p["nombre"], *p.get("aliases", [])]:
            sc = _score(raw, cand)
            if sc > best_score:
                best, best_score = p, sc
    if best and best_score >= umbral:
        return Person(raw, best["nombre"], best["titulo"], best["rol"], True)
    # Sin coincidencia: se deja visible para que el usuario la complete.
    return Person(raw, raw.strip(), "Invitado", "invitado", False)
