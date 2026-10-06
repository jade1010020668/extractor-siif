"""Perfiles de modelo según la memoria RAM del equipo."""
from __future__ import annotations

import ctypes
import os
import subprocess
import sys
from dataclasses import dataclass


@dataclass(frozen=True)
class Perfil:
    clave: str
    nombre: str
    modelo: str
    num_ctx: int          # contexto del modelo (tokens)
    part_chars: int       # texto de transcripción por llamada
    descarga_gb: float
    ram_min_gb: int


PERFILES = {
    "8gb": Perfil("8gb", "Equipo con 8 GB de RAM", "qwen2.5:7b-instruct",
                  num_ctx=8192, part_chars=7000, descarga_gb=4.7, ram_min_gb=8),
    "16gb": Perfil("16gb", "Equipo con 16 GB de RAM o más", "qwen2.5:14b-instruct",
                   num_ctx=16384, part_chars=12000, descarga_gb=9.0, ram_min_gb=16),
}


def ram_gb() -> float | None:
    """Memoria física total en GB (None si no se puede leer)."""
    try:
        if sys.platform == "win32":
            class MEMORYSTATUSEX(ctypes.Structure):
                _fields_ = [("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong),
                            ("ullTotalPhys", ctypes.c_ulonglong), ("ullAvailPhys", ctypes.c_ulonglong),
                            ("ullTotalPageFile", ctypes.c_ulonglong), ("ullAvailPageFile", ctypes.c_ulonglong),
                            ("ullTotalVirtual", ctypes.c_ulonglong), ("ullAvailVirtual", ctypes.c_ulonglong),
                            ("ullAvailExtendedVirtual", ctypes.c_ulonglong)]
            st = MEMORYSTATUSEX()
            st.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
            ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(st))
            return st.ullTotalPhys / 1024**3
        if sys.platform == "darwin":
            out = subprocess.run(["sysctl", "-n", "hw.memsize"], capture_output=True, text=True, check=True)
            return int(out.stdout) / 1024**3
        if os.path.exists("/proc/meminfo"):
            with open("/proc/meminfo") as f:
                for line in f:
                    if line.startswith("MemTotal:"):
                        return int(line.split()[1]) / 1024**2
    except Exception:                                # noqa: BLE001
        return None
    return None


def recomendado(ram: float | None) -> str:
    # Windows reporta algo menos de lo nominal (p. ej. 15,7 GB en un equipo de 16 GB)
    return "16gb" if ram is not None and ram >= 15 else "8gb"
