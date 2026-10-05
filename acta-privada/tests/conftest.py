import io
import json
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import pytest
from docx import Document

ROSTER = {
    "ciudad": "Bogotá D.C.",
    "personas": [
        {"nombre": "Ana María Pérez Gómez", "aliases": ["Ana Perez Gomez"], "rol": "presidente", "titulo": "Comisionada Presidente"},
        {"nombre": "Carlos Andrés Rojas Lema", "aliases": ["Carlos Rojas"], "rol": "comisionado", "titulo": "Comisionado"},
        {"nombre": "Lucía Fernanda Mora Díaz", "aliases": ["Lucia Mora"], "rol": "comisionado", "titulo": "Comisionada"},
        {"nombre": "Marta Elena Suárez Ríos", "aliases": ["Relatoria"], "rol": "relatoria", "titulo": "Profesional Universitaria"},
    ],
}

# Reunión 100 % ficticia, con el mismo patrón de exportación de Teams.
TRANSCRIPT_LINES = [
    "Sala Plena Virtual del 3 de marzo de 2027, a las 0930 a.m.-20270303_093001-Grabación de la reunión",
    "3 de marzo de 2027, 11:00a.m.",
    "15 min 30 s",
    "",
    "Relatoria ha iniciado la transcripción",
    "Relatoria   0:03",
    "Tenemos el orden del día: aprobación del orden del día y proposiciones y varios.",
    "Ana Perez Gomez   0:15",
    "Gracias, de acuerdo con el orden del día.",
    "Carlos Rojas   0:19",
    "De acuerdo con el orden del día.",
    "Lucia Mora   0:23",
    "De acuerdo, aprobado.",
    "Relatoria   0:30",
    "Entonces pasaríamos a abordar temas de proposiciones y varios.",
    "Ana Perez Gomez   0:40",
    "Quiero informar que se expidió la resolución 55 de 2027 sobre archivo documental.",
    "Carlos Rojas   2:10",
    "Considero que debemos revisar su impacto en cada despacho antes de actuar.",
    "Lucia Mora   3:00",
    "Estoy de acuerdo, que el área jurídica emita un concepto.",
    "Ana Perez Gomez   4:00",
    "Se aprueba entonces solicitar el concepto jurídico.",
    "Relatoria ha detenido la transcripción",
]


@pytest.fixture(autouse=True)
def _datos_aislados(tmp_path, monkeypatch):
    """Las pruebas nunca leen ni escriben la carpeta de datos real del usuario."""
    monkeypatch.setenv("ACTA_DATA_DIR", str(tmp_path / "datos"))


@pytest.fixture
def roster():
    return json.loads(json.dumps(ROSTER))


@pytest.fixture
def transcript_docx() -> bytes:
    d = Document()
    for line in TRANSCRIPT_LINES:
        d.add_paragraph(line)
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


class _FakeOllama(BaseHTTPRequestHandler):
    log: list = []
    modo = "ok"

    def log_message(self, *a):
        pass

    def _send(self, obj, code=200):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        self._send({"models": [{"name": "modelo-falso:latest"}]})

    def do_POST(self):
        req = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
        if self.path == "/api/pull":
            lines = [{"status": "pulling manifest"},
                     {"status": "downloading", "total": 100, "completed": 50},
                     {"status": "downloading", "total": 100, "completed": 100},
                     {"status": "success"}]
            body = "".join(json.dumps(x) + "\n" for x in lines).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/x-ndjson")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            _FakeOllama.log.append(("pull", req["model"]))
            return
        system = req["messages"][0]["content"]
        user = req["messages"][1]["content"]
        _FakeOllama.log.append((system[:40], user))
        if _FakeOllama.modo == "basura":
            return self._send({"message": {"content": "esto no es json"}})
        if "Identifica" in system:      # detección de temas
            out = {"orden_del_dia": ["Aprobación del Orden del Día", "Proposiciones y Varios"],
                   "temas": [{"punto": 2, "titulo": "Resolución 55 de 2027", "inicio": 4}]}
        elif "Tipos de párrafo" in system:   # redacción
            out = {"parrafos": [
                {"tipo": "presentacion", "texto": "La Comisionada Presidente Ana María Pérez Gómez, informa que se expidió la Resolución 55 de 2027."},
                {"tipo": "discusion", "texto": "El Comisionado Carlos Andrés Rojas Lema, manifiesta que debe revisarse su impacto."}]}
        else:                            # consolidación
            out = {"pretension": "La Presidente somete a consideración la Resolución 55 de 2027.",
                   "solicitud": ["Los Comisionados solicitan un concepto jurídico."],
                   "decision": "Los Comisionados deciden por unanimidad solicitar el concepto jurídico."}
        self._send({"message": {"content": json.dumps(out)}})


@pytest.fixture
def fake_ollama():
    _FakeOllama.log = []
    _FakeOllama.modo = "ok"
    srv = HTTPServer(("127.0.0.1", 0), _FakeOllama)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{srv.server_port}", _FakeOllama
    srv.shutdown()
