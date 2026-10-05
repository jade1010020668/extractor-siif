from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Segment:
    idx: int
    speaker: str          # nombre tal como aparece en la transcripción
    seconds: int
    text: str


@dataclass
class Transcript:
    titulo: str = ""
    fecha: tuple[int, int, int] | None = None   # (año, mes, día)
    hora_inicio: tuple[int, int] | None = None  # (hora 0-23, minuto)
    duracion_seg: int | None = None
    segments: list[Segment] = field(default_factory=list)


@dataclass
class Person:
    raw: str              # etiqueta en la transcripción
    nombre: str           # nombre canónico
    titulo: str           # p. ej. "Comisionada Presidente"
    rol: str              # presidente | comisionado | relatoria | invitado
    reconocido: bool = True

    @property
    def articulo(self) -> str:
        first = (self.titulo.split() or [""])[0].lower()
        return "La" if first.endswith("a") else "El"

    @property
    def tratamiento(self) -> str:
        """'La Comisionada Presidente Ana María Pérez Gómez'"""
        return f"{self.articulo} {self.titulo} {self.nombre}".replace("  ", " ")


@dataclass
class Tema:
    punto: int                       # punto del orden del día (1-based)
    titulo: str
    inicio: int                      # índice de segmento (incl.)
    fin: int                         # índice de segmento (incl.)
    pretension: str = ""
    presentacion: list[str] = field(default_factory=list)
    discusion: list[str] = field(default_factory=list)
    solicitud: list[str] = field(default_factory=list)
    decision: str = ""
    cierre: list[str] = field(default_factory=list)


@dataclass
class ActaData:
    numero: str = ""
    tipo_sesion: str = "SESIÓN ORDINARIA"
    fecha: tuple[int, int, int] | None = None
    ciudad: str = "Bogotá D.C."
    hora_inicio: str = ""
    hora_fin: str = ""
    presidente: Person | None = None
    comisionados: list[Person] = field(default_factory=list)
    relatoria: Person | None = None
    invitados: str = "No hubo invitados para la presente Sesión."
    orden_del_dia: list[str] = field(default_factory=list)
    apertura: str = ""
    temas: list[Tema] = field(default_factory=list)
    nota_firmas: str = ""
    avisos: list[str] = field(default_factory=list)   # puntos a verificar
