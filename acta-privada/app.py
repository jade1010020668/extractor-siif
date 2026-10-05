"""Interfaz local (Streamlit). Las transcripciones y actas solo viven en memoria."""
from __future__ import annotations

import pandas as pd
import streamlit as st

from acta_privada.docx_builder import build_docx, plantilla_activa
from acta_privada.llm import DEFAULT_HOST, LLMUnavailable, OllamaClient
from acta_privada.models import Person
from acta_privada.paths import (data_dir, escribir_json, guardar_ajustes, leer_ajustes,
                                user_config, user_plantilla)
from acta_privada.perfiles import PERFILES, ram_gb, recomendado
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


# ───────────────────────── barra lateral: motor ─────────────────────────
def panel_motor() -> OllamaClient | None:
    st.header("Motor de redacción")
    modo = st.radio("Modo", ["IA local (Ollama)", "Solo reglas (sin IA)"],
                    help="Con IA se redacta en tercera persona. Sin IA se obtiene un borrador "
                         "extractivo marcado [REVISAR].")
    if not modo.startswith("IA"):
        return None

    try:
        base = OllamaClient(host=DEFAULT_HOST)
        instalados = base.available_models()
    except (LLMUnavailable, PrivacyError) as e:
        st.error(f"{e}\n\nUse «Solo reglas» o reinicie la aplicación.")
        return None

    ram = ram_gb()
    rec = recomendado(ram)
    st.caption(f"Memoria de este equipo: **{ram:.0f} GB**" if ram else "Memoria del equipo: desconocida")
    claves = list(PERFILES)
    guardado = leer_ajustes().get("perfil", rec)

    def etiqueta(k: str) -> str:
        p = PERFILES[k]
        ok = p.modelo in instalados
        return f"{p.nombre}{' · recomendado' if k == rec else ''} {'✅' if ok else '⬇️'}"

    clave = st.radio("Modelo", claves, index=claves.index(guardado) if guardado in claves else 0,
                     format_func=etiqueta)
    if clave != leer_ajustes().get("perfil"):
        guardar_ajustes(perfil=clave)
    perfil = PERFILES[clave]
    if ram and ram < perfil.ram_min_gb - 1:
        st.warning(f"Este modelo necesita {perfil.ram_min_gb} GB de RAM; en este equipo puede ser muy lento.")

    if perfil.modelo not in instalados:
        st.info(f"Falta descargar **{perfil.modelo}** ({perfil.descarga_gb:.1f} GB). "
                "Solo se hace una vez; después funciona sin internet.")
        if st.button(f"⬇️ Descargar modelo ({perfil.descarga_gb:.1f} GB)", type="primary",
                     use_container_width=True):
            bar = st.progress(0.0, "Conectando…")
            try:
                base.pull(perfil.modelo, lambda f, s: bar.progress(f if f is not None else 0.0,
                                                                   f"{s} {f:.0%}" if f else s))
                st.success("Modelo listo.")
                st.rerun()
            except LLMUnavailable as e:
                st.error(str(e))
        return None

    llm = OllamaClient.desde_perfil(perfil)
    st.success(f"IA local lista · {perfil.modelo}")
    return llm


# ─────────────────────────── pestaña: configuración ───────────────────────────
def pestaña_configuracion():
    st.caption(f"Se guarda solo en este equipo, en `{data_dir()}`.")
    roster = load_json("roster")
    textos = load_json("textos")

    st.subheader("Participantes habituales")
    st.caption("«Alias» = cómo aparece el nombre en la transcripción de Teams (separe varios con «;»). "
               "El cargo se escribe tal como debe salir en el acta (p. ej. «Comisionada Presidente»).")
    ciudad = st.text_input("Ciudad", roster.get("ciudad", "Bogotá D.C."))
    df = pd.DataFrame([{"Nombre completo": p["nombre"], "Alias": "; ".join(p.get("aliases", [])),
                        "Cargo": p["titulo"], "Rol": p["rol"]} for p in roster.get("personas", [])]
                      or [{"Nombre completo": "", "Alias": "", "Cargo": "", "Rol": "comisionado"}])
    df = st.data_editor(df, num_rows="dynamic", hide_index=True, use_container_width=True,
                        column_config={"Rol": st.column_config.SelectboxColumn(options=ROLES, required=True)})

    st.subheader("Textos fijos del acta")
    campos = {
        "tipo_sesion": "Tipo de sesión",
        "parrafo_virtual": "Párrafo de participación virtual",
        "aprobacion_unanime": "Aprobación del orden del día (unanimidad)",
        "aprobacion_sin_unanimidad": "Aprobación del orden del día (sin unanimidad)",
        "sin_invitados": "Texto cuando no hay invitados",
        "nota_firmas": "Nota de firmas",
    }
    nuevos = {k: st.text_area(v, textos.get(k, ""), height=80 if len(textos.get(k, "")) > 90 else 68)
              for k, v in campos.items()}
    orden = st.text_area("Orden del día por defecto (un punto por línea)",
                         "\n".join(textos.get("orden_del_dia_por_defecto", [])), height=80)

    if st.button("💾 Guardar configuración", type="primary"):
        personas = [{"nombre": r["Nombre completo"].strip(),
                     "aliases": [a.strip() for a in str(r["Alias"] or "").split(";") if a.strip()],
                     "rol": r["Rol"] or "comisionado", "titulo": (r["Cargo"] or "").strip()}
                    for _, r in df.iterrows() if str(r["Nombre completo"] or "").strip()]
        escribir_json(user_config("roster"), {"ciudad": ciudad, "personas": personas})
        nuevos["orden_del_dia_por_defecto"] = [x.strip() for x in orden.splitlines() if x.strip()]
        escribir_json(user_config("textos"), textos | nuevos)
        st.success("Configuración guardada.")

    st.subheader("Plantilla (logo, encabezado y pie de página)")
    actual = plantilla_activa()
    st.caption(f"Plantilla en uso: `{actual.name}`" if actual else
               "Sin plantilla: se genera un documento con formato equivalente pero sin logo.")
    st.caption("Suba un acta anterior: solo se conservan su encabezado, pie, márgenes y estilos; "
               "el texto y los metadatos se eliminan en cada acta nueva.")
    up = st.file_uploader("Plantilla .docx", type=["docx"], key="plantilla_up")
    c1, c2 = st.columns(2)
    if up and c1.button("Usar esta plantilla"):
        user_plantilla().parent.mkdir(parents=True, exist_ok=True)
        user_plantilla().write_bytes(up.getvalue())
        st.success("Plantilla guardada.")
    if user_plantilla().exists() and c2.button("Quitar plantilla"):
        user_plantilla().unlink()
        st.success("Plantilla eliminada.")


# ─────────────────────────── pestaña: generar acta ───────────────────────────
def pestaña_generar(llm: OllamaClient | None):
    roster = load_json("roster")
    textos = load_json("textos")

    up = st.file_uploader("1 · Transcripción de la reunión (.docx o .txt)", type=["docx", "txt"])
    if up is None:
        st.info("Suba la transcripción para comenzar. Si es la primera vez, revise antes la pestaña "
                "«Configuración».")
        return
    try:
        tr = read_any(up.name, up.getvalue())
    except Exception as e:                                # noqa: BLE001
        st.error(f"No se pudo leer la transcripción: {e}")
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("Intervenciones", len(tr.segments))
    c2.metric("Fecha", "-" if not tr.fecha else f"{tr.fecha[2]:02d}/{tr.fecha[1]:02d}/{tr.fecha[0]}")
    c3.metric("Duración", "-" if not tr.duracion_seg else f"{tr.duracion_seg // 60} min")

    st.subheader("2 · Participantes")
    st.caption("Se reconocen con los participantes de «Configuración». Corrija aquí si hace falta.")
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

    if st.button("3 · Generar acta", type="primary"):
        if llm is None:
            st.caption("Generando en modo «Solo reglas».")
        bar = st.progress(0.0, "Iniciando… (con IA puede tardar varios minutos)")
        st.session_state["acta"] = build_acta(
            tr, roster, numero=numero, llm=llm, textos=textos, people=people,
            progress=lambda f, m: bar.progress(min(f, 1.0), m))
        bar.empty()
        for k in [k for k in st.session_state if k.startswith("ed_")]:
            del st.session_state[k]

    acta = st.session_state.get("acta")
    if not acta:
        return

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
    acta.apertura = st.text_area("Punto 1 — Aprobación del orden del día", acta.apertura,
                                 key="ed_ap", height=140)
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


# ──────────────────────────────── página ────────────────────────────────
st.title("🔒 Generador de actas — 100 % local")
st.caption("La transcripción se procesa solo en este equipo. Nada se envía a internet ni a la nube; "
           "no se guarda en disco y se descarta al cerrar o pulsar «Borrar todo».")
with st.sidebar:
    motor = panel_motor()
    st.divider()
    st.button("🗑️ Borrar todo de la sesión", on_click=borrar_todo, use_container_width=True)

tab_acta, tab_conf = st.tabs(["📝 Generar acta", "⚙️ Configuración"])
with tab_acta:
    pestaña_generar(motor)
with tab_conf:
    pestaña_configuracion()
