"""Garantías de privacidad: el contenido de la reunión nunca sale del equipo.

- El único servicio de red que usa la app es Ollama, y solo se acepta si
  apunta a este mismo equipo (loopback). Para usar un servidor Ollama de la
  red interna hay que autorizarlo explícitamente (ACTA_PERMITIR_RED_LOCAL=1).
- Las llamadas HTTP ignoran proxies del sistema (HTTP(S)_PROXY): un proxy
  corporativo o en la nube NO debe ver el tráfico hacia Ollama.
"""
from __future__ import annotations

import ipaddress
import os
import socket
from urllib.parse import urlparse


class PrivacyError(RuntimeError):
    pass


def _is_loopback(host: str) -> bool:
    if host in ("localhost", "ip6-localhost"):
        return True
    try:
        return ipaddress.ip_address(host).is_loopback
    except ValueError:
        return False


def _is_private(host: str) -> bool:
    try:
        ip = ipaddress.ip_address(host)
    except ValueError:
        try:
            ip = ipaddress.ip_address(socket.gethostbyname(host))
        except OSError:
            return False
    return ip.is_private or ip.is_loopback or ip.is_link_local


def validate_endpoint(url: str) -> str:
    """Devuelve la URL normalizada o lanza PrivacyError si no es local."""
    if "://" not in url:
        url = "http://" + url
    parsed = urlparse(url)
    host = parsed.hostname or ""
    if parsed.scheme not in ("http", "https") or not host:
        raise PrivacyError(f"URL de Ollama no válida: {url!r}")
    if _is_loopback(host):
        return url.rstrip("/")
    if os.environ.get("ACTA_PERMITIR_RED_LOCAL") == "1" and _is_private(host):
        return url.rstrip("/")
    raise PrivacyError(
        f"Bloqueado: {host!r} no es este equipo. Para proteger la información "
        "solo se permite Ollama en localhost (127.0.0.1). Si tiene un servidor "
        "Ollama propio en su red interna, defina ACTA_PERMITIR_RED_LOCAL=1."
    )
