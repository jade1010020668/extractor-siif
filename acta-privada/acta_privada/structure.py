"""Transcripción -> datos del acta.

Dos motores, mismo resultado (ActaData):
  * Con IA local (Ollama): identifica temas y redacta en tercera persona.
  * Sin IA (reglas): borrador extractivo, marcado [REVISAR], 100 % determinista.

Las reglas siempre hacen lo "duro": participantes, cargos, horas, fecha,
unanimidad del orden del día y formato. La IA solo redacta narrativa.
"""
from __future__ import annotations

import re
from typing import Callable

from .llm import LLMUnavailable, OllamaClient
from .models import ActaData, Person, Segment, Tema, Transcript
from .roster import load_json, match_person

Progress = Callable[[float, str], None]
PART_CHARS = 12_000          # texto de transcripción por llamada al modelo
INDEX_CHARS = 12_000         # índice condensado por llamada (detección de temas)
REVISAR = "[REVISAR]"


# ───────────────────────── utilidades de reglas ─────────────────────────
def format_hora(h: int, m: int) -> str:
    return f"{(h % 12) or 12}:{m:02d} {'p.m.' if h >= 12 else 'a.m.'}"


def compute_horas(tr: Transcript) -> tuple[str, str]:
    if not tr.hora_inicio:
        return "", ""
    h, m = tr.hora_inicio
    inicio = format_hora(h, m)
    dur = tr.duracion_seg or (tr.segments[-1].seconds if tr.segments else 0)
    total = h * 60 + m + round(dur / 60)
    return inicio, format_hora((total // 60) % 24, total % 60)


def resolve_people(tr: Transcript, roster: dict) -> dict[str, Person]:
    people: dict[str, Person] = {}
    for s in tr.segments:
        if s.speaker not in people:
            people[s.speaker] = match_person(s.speaker, roster)
    return people


_AFIRMA = re.compile(r"de acuerdo|apruebo|aprobado|conforme|a favor", re.I)
_TRANSICION = re.compile(
    r"(pasar[ií]amos|pasamos|abordar|siguiente punto|continuamos).*(proposiciones|varios|punto)", re.I)


def split_apertura(tr: Transcript) -> int:
    """Índice del primer segmento posterior a la aprobación del orden del día
    (solo se usa en modo reglas; con IA lo decide la detección de temas)."""
    for s in tr.segments[:40]:
        if s.idx > 0 and _TRANSICION.search(s.text):
            return s.idx + 1
    return 0


def aprobacion_orden_del_dia(tr: Transcript, fin_apertura: int, people: dict[str, Person],
                             textos: dict) -> tuple[str, str | None]:
    """Texto de aprobación + aviso. Unanimidad solo si TODOS los comisionados
    que intervinieron en la sala expresaron conformidad en la apertura."""
    comisionados = {p.nombre for p in people.values() if p.rol in ("presidente", "comisionado")}
    afirmaron = {
        people[s.speaker].nombre for s in tr.segments[:fin_apertura]
        if people[s.speaker].nombre in comisionados and _AFIRMA.search(s.text)
    }
    if comisionados and afirmaron == comisionados:
        return textos["aprobacion_unanime"], None
    faltan = sorted(comisionados - afirmaron)
    return textos["aprobacion_sin_unanimidad"], (
        "No se pudo confirmar la unanimidad en la aprobación del orden del día"
        + (f" (sin conformidad expresa de: {', '.join(faltan)})." if faltan else "."))


def _label(s: Segment, people: dict[str, Person]) -> str:
    return people[s.speaker].tratamiento


def _merge_turns(segs: list[Segment]) -> list[tuple[str, str]]:
    out: list[tuple[str, str]] = []
    for s in segs:
        if out and out[-1][0] == s.speaker:
            out[-1] = (s.speaker, out[-1][1] + " " + s.text)
        else:
            out.append((s.speaker, s.text))
    return out


def _chunks(items: list[str], budget: int) -> list[list[str]]:
    parts, cur, size = [], [], 0
    for it in items:
        if cur and size + len(it) > budget:
            parts.append(cur)
            cur, size = [], 0
        cur.append(it)
        size += len(it) + 1
    if cur:
        parts.append(cur)
    return parts


# ───────────────────────────── motor con IA ─────────────────────────────
_TEMAS_SCHEMA = {
    "type": "object",
    "properties": {
        "orden_del_dia": {"type": "array", "items": {"type": "string"}},
        "temas": {"type": "array", "items": {"type": "object", "properties": {
            "punto": {"type": "integer"}, "titulo": {"type": "string"},
            "inicio": {"type": "integer"}, "continua": {"type": "boolean"}},
            "required": ["punto", "titulo", "inicio"]}},
    },
    "required": ["orden_del_dia", "temas"],
}
_PARRAFOS_SCHEMA = {
    "type": "object",
    "properties": {"parrafos": {"type": "array", "items": {"type": "object", "properties": {
        "tipo": {"type": "string", "enum": ["presentacion", "discusion", "cierre"]},
        "texto": {"type": "string"}}, "required": ["tipo", "texto"]}}},
    "required": ["parrafos"],
}
_CONSOLIDA_SCHEMA = {
    "type": "object",
    "properties": {"pretension": {"type": "string"},
                   "solicitud": {"type": "array", "items": {"type": "string"}},
                   "decision": {"type": "string"}},
    "required": ["pretension", "solicitud", "decision"],
}

_SYS_TEMAS = (
    "Analizas la transcripción de la sesión de un órgano colegiado. Identifica: "
    "(1) el orden del día (puntos que la presidencia o relatoría enumera; si no se "
    "lee completo, deduce los puntos tratados); (2) los temas tratados, indicando el "
    "número [n] del segmento donde cada uno COMIENZA. 'Aprobación del Orden del Día' "
    "no es un tema: no lo incluyas en 'temas'. Si dentro de 'Proposiciones y Varios' "
    "se trata un asunto concreto, ese asunto es un tema cuyo 'punto' es el número de "
    "'Proposiciones y Varios'. Si un tema ya abierto continúa al inicio de este "
    "fragmento, marca 'continua': true. No inventes nada. Responde solo JSON."
)
_SYS_REDACCION = (
    "Eres relator/a de actas de una Sala Plena. Redacta SOLO con lo dicho en la "
    "transcripción; no inventes hechos, cifras, normas, fechas ni nombres. Estilo: "
    "tercera persona, tiempo presente, tono jurídico-administrativo formal, p. ej. "
    "'La Comisionada Presidente Ana María Pérez Gómez, informa a la Sala Plena que ...', "
    "'Posteriormente, el Comisionado Carlos Rojas Lema, manifiesta que ...', "
    "'Por último, ...'. Usa EXACTAMENTE los tratamientos y nombres de la lista de "
    "participantes. Un párrafo por intervención o grupo coherente de intervenciones; "
    "resume lo sustancial y omite saludos, agradecimientos y muletillas. Corrige "
    "errores evidentes de la transcripción automática, pero si un número de norma, "
    "artículo, fecha o cifra parece dudoso escribe [VERIFICAR] junto a él. "
    "Tipos de párrafo: 'presentacion' (quien introduce el tema), 'discusion' "
    "(intervenciones), 'cierre' (cierre del punto o ausencia de más temas). "
    "Responde solo JSON."
)
_SYS_CONSOLIDA = (
    "Eres relator/a de actas. A partir del relato de un punto de la sesión, escribe: "
    "'pretension' (una frase: quién somete a consideración de la Sala qué asunto, en "
    "tercera persona formal), 'solicitud' (lista de párrafos con lo que los Comisionados "
    "solicitan a áreas o despachos, si se pidió algo; si no, lista vacía) y 'decision' "
    "(lo que la Sala decide o aprueba, si hubo decisión explícita; si no, cadena vacía). "
    "No inventes: si no hay solicitud o decisión explícita, déjalas vacías. Responde solo JSON."
)


def _participantes_txt(people: dict[str, Person]) -> str:
    uniq = {p.nombre: p for p in people.values()}
    return "\n".join(f"- {p.tratamiento}" for p in uniq.values())


def _detectar_temas(llm: OllamaClient, tr: Transcript, people: dict[str, Person],
                    default_orden: list[str]) -> tuple[list[str], list[dict]]:
    lines = [f"[{s.idx}] {_label(s, people)}: {s.text[:220]}" for s in tr.segments]
    orden: list[str] = []
    temas: list[dict] = []
    for chunk in _chunks(lines, INDEX_CHARS):
        ctx = ""
        if temas:
            ult = temas[-1]
            ctx = (f"Orden del día ya identificado: {orden}\n"
                   f"Último tema abierto: punto {ult['punto']}, «{ult['titulo']}».\n\n")
        out = llm.chat_json(_SYS_TEMAS, ctx + "Segmentos:\n" + "\n".join(chunk), _TEMAS_SCHEMA)
        for item in out.get("orden_del_dia", []):
            item = str(item).strip()
            if item and item.lower() not in (o.lower() for o in orden):
                orden.append(item)
        for t in out.get("temas", []):
            if t.get("continua") and temas:
                continue
            try:
                temas.append({"punto": int(t["punto"]), "titulo": str(t["titulo"]).strip(),
                              "inicio": int(t["inicio"])})
            except (KeyError, ValueError, TypeError):
                continue
    return orden or list(default_orden), temas


def _normalizar_temas(raw: list[dict], orden: list[str], n_seg: int, fin_apertura_min: int) -> list[Tema]:
    raw = sorted((t for t in raw if 0 <= t["inicio"] < n_seg), key=lambda t: t["inicio"])
    dedup: list[dict] = []
    for t in raw:
        if not dedup or t["inicio"] > dedup[-1]["inicio"]:
            dedup.append(t)
    out: list[Tema] = []
    for i, t in enumerate(dedup):
        fin = (dedup[i + 1]["inicio"] - 1) if i + 1 < len(dedup) else n_seg - 1
        punto = t["punto"] if 1 <= t["punto"] <= len(orden) else len(orden)
        out.append(Tema(punto=punto, titulo=t["titulo"], inicio=t["inicio"], fin=fin))
    return out


def _redactar_tema(llm: OllamaClient, tr: Transcript, tema: Tema, people: dict[str, Person],
                   lista: str) -> None:
    segs = tr.segments[tema.inicio: tema.fin + 1]
    lines = [f"{_label(s, people)}: {s.text}" for s in segs]
    parrafos: list[tuple[str, str]] = []
    for part in _chunks(lines, PART_CHARS):
        user = (f"Participantes:\n{lista}\n\nTema: {tema.titulo}\n\nTranscripción:\n" + "\n".join(part))
        out = llm.chat_json(_SYS_REDACCION, user, _PARRAFOS_SCHEMA)
        for p in out.get("parrafos", []):
            texto = str(p.get("texto", "")).strip()
            if texto:
                parrafos.append((p.get("tipo", "discusion"), texto))
    tema.presentacion = [t for k, t in parrafos if k == "presentacion"]
    tema.discusion = [t for k, t in parrafos if k == "discusion"]
    tema.cierre = [t for k, t in parrafos if k == "cierre"]
    relato = "\n\n".join(t for _, t in parrafos)
    if len(relato) > PART_CHARS * 2:               # conserva inicio y final
        relato = relato[:PART_CHARS] + "\n[...]\n" + relato[-PART_CHARS:]
    out = llm.chat_json(_SYS_CONSOLIDA, f"Participantes:\n{lista}\n\nTema: {tema.titulo}\n\nRelato:\n{relato}",
                        _CONSOLIDA_SCHEMA)
    tema.pretension = str(out.get("pretension", "")).strip()
    tema.solicitud = [str(x).strip() for x in out.get("solicitud", []) if str(x).strip()]
    tema.decision = str(out.get("decision", "")).strip()


# ──────────────────────────── motor por reglas ───────────────────────────
_CORTESIA = re.compile(r"^(muchas |muchísimas )?gracias|^de acuerdo|^claro que s", re.I)


def _borrador_reglas(tr: Transcript, people: dict[str, Person], inicio: int) -> list[str]:
    out = []
    for spk, texto in _merge_turns(tr.segments[inicio:]):
        if len(texto.split()) < 8 and _CORTESIA.search(texto):
            continue
        out.append(f"{REVISAR} {people[spk].tratamiento}: {texto}")
    return out


# ──────────────────────────────── API ───────────────────────────────────
def build_acta(
    tr: Transcript,
    roster: dict,
    *,
    numero: str = "",
    llm: OllamaClient | None = None,
    textos: dict | None = None,
    people: dict[str, Person] | None = None,
    progress: Progress | None = None,
) -> ActaData:
    textos = textos or load_json("textos")
    people = people or resolve_people(tr, roster)
    say = progress or (lambda f, m: None)
    acta = ActaData(numero=numero, tipo_sesion=textos.get("tipo_sesion", "SESIÓN ORDINARIA"),
                    fecha=tr.fecha, ciudad=roster.get("ciudad", "Bogotá D.C."),
                    invitados=textos.get("sin_invitados", ""),
                    nota_firmas=textos.get("nota_firmas", ""))
    acta.hora_inicio, acta.hora_fin = compute_horas(tr)
    if not acta.hora_inicio:
        acta.avisos.append("No se pudo leer la hora de inicio de la transcripción; complétela.")

    uniq = {p.nombre: p for p in people.values()}.values()
    pres = [p for p in uniq if p.rol == "presidente"]
    acta.presidente = pres[0] if pres else None
    orden_padron = {p["nombre"]: i for i, p in enumerate(roster.get("personas", []))}
    acta.comisionados = sorted((p for p in uniq if p.rol == "comisionado"),
                               key=lambda p: orden_padron.get(p.nombre, 999))
    rel = [p for p in uniq if p.rol == "relatoria"]
    acta.relatoria = rel[0] if rel else None
    ext = [p for p in uniq if p.rol == "invitado"]
    if ext:
        acta.invitados = "; ".join(f"{p.nombre}" for p in ext)
    for p in uniq:
        if not p.reconocido:
            acta.avisos.append(f"Hablante no reconocido en el padrón: «{p.raw}». Asigne nombre y cargo.")
    if not acta.presidente:
        acta.avisos.append("No hay Presidente identificado entre los hablantes.")

    default_orden = textos.get("orden_del_dia_por_defecto", ["Aprobación del Orden del Día", "Proposiciones y Varios"])
    n = len(tr.segments)

    if llm is not None:
        try:
            say(0.05, "Identificando orden del día y temas…")
            orden, raw = _detectar_temas(llm, tr, people, default_orden)
            acta.orden_del_dia = orden
            acta.temas = _normalizar_temas(raw, orden, n, 0)
            if not acta.temas:
                raise LLMUnavailable("el modelo no identificó temas")
            fin_apertura = acta.temas[0].inicio
            lista = _participantes_txt(people)
            for i, tema in enumerate(acta.temas):
                say(0.1 + 0.85 * i / len(acta.temas), f"Redactando: {tema.titulo[:60]}…")
                _redactar_tema(llm, tr, tema, people, lista)
        except LLMUnavailable as exc:
            acta.avisos.append(f"La IA local falló ({exc}); se generó un borrador por reglas.")
            llm = None

    if llm is None:
        fin_apertura = split_apertura(tr)
        acta.orden_del_dia = list(default_orden)
        acta.temas = [Tema(punto=len(default_orden), titulo=f"{REVISAR} Tema tratado",
                           inicio=fin_apertura, fin=n - 1,
                           discusion=_borrador_reglas(tr, people, fin_apertura))]
        acta.avisos.append("Borrador por reglas: el texto de los temas es extractivo; edítelo antes de firmar.")

    acta.apertura, aviso = aprobacion_orden_del_dia(tr, fin_apertura, people, textos)
    acta.apertura = textos.get("parrafo_virtual", "") + "\n\n" + acta.apertura
    if aviso:
        acta.avisos.append(aviso)
    texto_total = " ".join([acta.apertura] + [x for t in acta.temas for x in
                           [t.pretension, t.decision, *t.presentacion, *t.discusion, *t.solicitud, *t.cierre]])
    if "[VERIFICAR]" in texto_total:
        acta.avisos.append(f"Hay {texto_total.count('[VERIFICAR]')} marca(s) [VERIFICAR] en el texto.")
    say(1.0, "Listo")
    return acta
