"""Interfaz local (Streamlit). Todo ocurre en memoria; no se escribe nada a disco."""
from __future__ import annotations

import pandas as pd
import streamlit as st

from acta_privada.docx_builder import build_docx
from acta_privada.llm import DEFAULT_HOST, DEFAULT_MODEL, LLMUnavailable, OllamaClient
from acta_privada.models import Person
from acta_privada.privacy import PrivacyError
from acta_privada.roster import load_json, match_person
from acta_privada.structure import build_acta
from acta_privada.transcript import read_any

st.set_page_config(page_title="Actas privadas", page_icon="🔒", layout="wide")
ROLES = ["presidente", "comisionado", "relatoria", "invitado"]


def _split(txt: str) -> list[str]:
    return [p.strip() for p in txt.split("\n\n") if p.strip()]


def borrar_todo():
    for k in list(st.session_state.keys()):
        del st.session_state[k]


st.title("🔒 Generador de actas — 100 % local")
st.caption("La transcripción se procesa solo en este equipo. Nada se envía a internet ni a la nube; "
           "no se guarda en disco y se descarta al cerrar o pulsar «Borrar todo».")

# ───────────── barra lateral: motor ─────────────
with st.sidebar:
    st.header("Motor de redacción")
    modo = st.radio("Modo", ["IA local (Ollama)", "Solo reglas (sin IA)"],
                    help="Con IA se redacta en tercera persona. Sin IA se obtiene un borrador "
                         "extractivo marcado [REVISAR].")
    llm = None
    if modo.startswith("IA"):
        host = st.text_input("Servidor Ollama", DEFAULT_HOST)
        modelo = st.text_input("Modelo", DEFAULT_MODEL)
        try:
            llm = OllamaClient(host=host, model=modelo)
            if llm.has_model():
                st.success(f"Ollama local listo · {modelo}")
            else:
                st.warning(f"Falta el modelo. En una terminal: `ollama pull {modelo}`")
                llm = None
        except PrivacyError as e:
            st.error(str(e))
        except LLMUnavailable as e:
            st.error(str(e))
    st.divider()
    st.button("🗑️ Borrar todo de la sesión", on_click=borrar_todo, use_container_width=True)

roster = load_json("roster")
textos = load_json("textos")

# ───────────── 1. cargar ─────────────
up = st.file_uploader("1 · Transcripción de la reunión (.docx o .txt)", type=["docx", "txt"])
if up is None:
    st.info("Suba la transcripción para comenzar.")
    st.stop()

try:
    tr = read_any(up.name, up.getvalue())
except Exception as e:                                    # noqa: BLE001
    st.error(f"No se pudo leer la transcripción: {e}")
    st.stop()

c1, c2, c3 = st.columns(3)
c1.metric("Intervenciones", len(tr.segments))
c2.metric("Fecha", "-" if not tr.fecha else f"{tr.fecha[2]:02d}/{tr.fecha[1]:02d}/{tr.fecha[0]}")
c3.metric("Duración", "-" if not tr.duracion_seg else f"{tr.duracion_seg // 60} min")

# ───────────── 2. participantes ─────────────
st.subheader("2 · Participantes")
st.caption("Se reconocen con el padrón local (config/roster.json). Complete los que falten.")
raws = list(dict.fromkeys(s.speaker for s in tr.segments))
pre = [match_person(r, roster) for r in raws]
df = pd.DataFrame({"En la transcripción": raws,
                   "Nombre en el acta": [p.nombre for p in pre],
                   "Cargo": [p.titulo for p in pre],
                   "Rol": [p.rol for p in pre]})
df = st.data_editor(df, hide_index=True, use_container_width=True, disabled=["En la transcripción"],
                    column_config={"Rol": st.column_config.SelectboxColumn(options=ROLES, required=True)})
people = {r["En la transcripción"]: Person(r["En la transcripción"], r["Nombre en el acta"],
                                           r["Cargo"], r["Rol"]) for _, r in df.iterrows()}

numero = st.text_input("Número de acta", "")

# ───────────── 3. generar ─────────────
if st.button("3 · Generar acta", type="primary"):
    bar = st.progress(0.0, "Iniciando…")
    st.session_state["acta"] = build_acta(
        tr, roster, numero=numero, llm=llm, textos=textos, people=people,
        progress=lambda f, m: bar.progress(min(f, 1.0), m))
    bar.empty()
    for k in [k for k in st.session_state if k.startswith("ed_")]:
        del st.session_state[k]

acta = st.session_state.get("acta")
if not acta:
    st.stop()

# ───────────── 4. revisar y descargar ─────────────
st.subheader("4 · Revisar y ajustar")
for av in acta.avisos:
    st.warning(av)

acta.numero = numero or acta.numero
a, b, c = st.columns(3)
acta.hora_inicio = a.text_input("Hora de inicio", acta.hora_inicio, key="ed_hi")
acta.hora_fin = b.text_input("Hora de finalización", acta.hora_fin, key="ed_hf")
acta.invitados = c.text_input("Invitados", acta.invitados, key="ed_inv")
acta.orden_del_dia = _split(st.text_area("Orden del día (un punto por párrafo)",
                                         "\n\n".join(acta.orden_del_dia), key="ed_od", height=110))
acta.apertura = st.text_area("Punto 1 — Aprobación del orden del día", acta.apertura, key="ed_ap", height=140)

for i, t in enumerate(acta.temas):
    with st.expander(f"Punto {t.punto} · {t.titulo}", expanded=True):
        t.titulo = st.text_input("Título del tema", t.titulo, key=f"ed_t{i}")
        t.pretension = st.text_area("Pretensión", t.pretension, key=f"ed_p{i}", height=80)
        t.presentacion = _split(st.text_area("Presentación del tema (párrafos separados por línea en blanco)",
                                             "\n\n".join(t.presentacion), key=f"ed_pr{i}", height=160))
        t.discusion = _split(st.text_area("Discusión sobre el tema", "\n\n".join(t.discusion),
                                          key=f"ed_d{i}", height=260))
        t.solicitud = _split(st.text_area("Solicitud", "\n\n".join(t.solicitud), key=f"ed_s{i}", height=100))
        t.decision = st.text_area("Decisión", t.decision, key=f"ed_de{i}", height=80)
        t.cierre = _split(st.text_area("Cierre", "\n\n".join(t.cierre), key=f"ed_c{i}", height=70))

pendientes = sum(x.count("[REVISAR]") + x.count("[VERIFICAR]") + x.count("[COMPLETAR]")
                 for t in acta.temas for x in [t.pretension, t.decision, *t.presentacion,
                                              *t.discusion, *t.solicitud, *t.cierre])
if pendientes:
    st.info(f"Quedan {pendientes} marca(s) [REVISAR]/[VERIFICAR] por resolver. Deje el texto final sin marcas.")

st.download_button("⬇️ Descargar acta (.docx)", data=build_docx(acta),
                   file_name=f"Acta_{acta.numero or 'borrador'}.docx",
                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                   type="primary")
st.caption("⚠️ El acta es un borrador asistido: debe ser revisada y firmada por las personas responsables.")
