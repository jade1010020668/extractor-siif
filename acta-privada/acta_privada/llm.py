"""Cliente mínimo de Ollama (solo biblioteca estándar, solo localhost)."""
from __future__ import annotations

import json
import os
import re
import urllib.error
import urllib.request
from dataclasses import dataclass, field

from .privacy import validate_endpoint

DEFAULT_HOST = os.environ.get("OLLAMA_HOST", "http://127.0.0.1:11434")
DEFAULT_MODEL = os.environ.get("ACTA_MODELO", "qwen2.5:14b-instruct")


class LLMUnavailable(RuntimeError):
    pass


# Opener sin proxies: nunca enrutar el tráfico local por un proxy externo.
_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def _extract_json(text: str) -> dict:
    text = text.strip()
    text = re.sub(r"^```(?:json)?|```$", "", text, flags=re.M).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start, end = text.find("{"), text.rfind("}")
        if start >= 0 and end > start:
            return json.loads(text[start : end + 1])
        raise


@dataclass
class OllamaClient:
    host: str = DEFAULT_HOST
    model: str = DEFAULT_MODEL
    num_ctx: int = 16384
    temperature: float = 0.2
    timeout: int = 900
    calls: int = field(default=0, init=False)

    def __post_init__(self):
        self.host = validate_endpoint(self.host)

    def _request(self, path: str, payload: dict | None = None) -> dict:
        data = json.dumps(payload).encode() if payload is not None else None
        req = urllib.request.Request(
            self.host + path,
            data=data,
            headers={"Content-Type": "application/json"},
            method="POST" if data is not None else "GET",
        )
        try:
            with _OPENER.open(req, timeout=self.timeout) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except (urllib.error.URLError, TimeoutError, ConnectionError) as exc:
            raise LLMUnavailable(
                f"No se pudo hablar con Ollama en {self.host} ({exc}). "
                "¿Está instalado y ejecutándose (ollama serve)?"
            ) from exc

    def available_models(self) -> list[str]:
        return [m["name"] for m in self._request("/api/tags").get("models", [])]

    def has_model(self) -> bool:
        names = self.available_models()
        return self.model in names or any(n.split(":")[0] == self.model for n in names)

    def chat_json(self, system: str, user: str, schema: dict | None = None) -> dict:
        """Pide una respuesta JSON; reintenta una vez si viene malformada."""
        last_err: Exception | None = None
        for attempt in range(2):
            payload = {
                "model": self.model,
                "stream": False,
                "format": schema if (schema and attempt == 0) else "json",
                "options": {"temperature": self.temperature, "num_ctx": self.num_ctx},
                "messages": [
                    {"role": "system", "content": system},
                    {"role": "user", "content": user},
                ],
            }
            self.calls += 1
            out = self._request("/api/chat", payload)
            try:
                return _extract_json(out["message"]["content"])
            except (KeyError, json.JSONDecodeError) as exc:
                last_err = exc
        raise LLMUnavailable(f"El modelo no devolvió JSON válido: {last_err}")
