import json
import socket
import sys
from pathlib import Path

import pytest

from acta_privada import paths
from acta_privada.__main__ import main as cli
from acta_privada.llm import OllamaClient
from acta_privada.perfiles import PERFILES, ram_gb, recomendado
from acta_privada.privacy import PrivacyError
from acta_privada.roster import load_json

FIXTURE = Path(__file__).parent / "fixtures" / "transcripcion_ficticia.txt"
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import launcher  # noqa: E402


@pytest.mark.parametrize("modelo", ["gpt-oss:120b-cloud", "deepseek-v3.1:671b-cloud", "algo:Cloud"])
def test_bloquea_modelos_en_la_nube(modelo):
    with pytest.raises(PrivacyError):
        OllamaClient(model=modelo)


def test_perfiles():
    assert PERFILES["8gb"].modelo == "qwen2.5:7b-instruct"
    assert PERFILES["16gb"].modelo == "qwen2.5:14b-instruct"
    assert recomendado(7.8) == "8gb" and recomendado(15.7) == "16gb" and recomendado(None) == "8gb"
    assert ram_gb() is None or ram_gb() > 0
    c = OllamaClient.desde_perfil(PERFILES["16gb"])
    assert (c.num_ctx, c.part_chars) == (16384, 12000)


def test_descarga_de_modelo_con_progreso(fake_ollama):
    url, srv = fake_ollama
    vistos = []
    OllamaClient(host=url, model="qwen2.5:7b-instruct").pull(progress=lambda f, s: vistos.append(f))
    assert ("pull", "qwen2.5:7b-instruct") in srv.log
    assert 0.5 in vistos and 1.0 in vistos


def test_configuracion_del_usuario_tiene_prioridad(tmp_path):
    paths.escribir_json(paths.user_config("roster"), {"ciudad": "Cali", "personas": []})
    assert load_json("roster")["ciudad"] == "Cali"
    paths.guardar_ajustes(perfil="16gb")
    assert paths.leer_ajustes() == {"perfil": "16gb"}
    assert str(paths.data_dir()).startswith(str(tmp_path))


def test_cli_sin_ia(tmp_path):
    out = tmp_path / "acta.docx"
    assert cli([str(FIXTURE), "-o", str(out), "--numero", "9", "--sin-ia"]) == 0
    assert out.stat().st_size > 5000


def test_cli_con_ia(tmp_path, fake_ollama):
    url, _ = fake_ollama
    out = tmp_path / "acta.docx"
    assert cli([str(FIXTURE), "-o", str(out), "--host", url, "--modelo", "modelo-falso"]) == 0


def test_launcher_puerto_libre():
    port = launcher.puerto_libre(18500)
    with socket.socket() as s:
        s.bind(("127.0.0.1", port))          # realmente estaba libre
    assert not launcher.responde("http://127.0.0.1:9/")


def test_instalar_modelo_descarga_y_guarda_perfil(fake_ollama):
    import instalar_modelo
    url, srv = fake_ollama
    msgs = []
    assert instalar_modelo.descargar("qwen2.5:7b-instruct", url, msgs.append)
    assert ("pull", "qwen2.5:7b-instruct") in srv.log
    assert any("100%" in m for m in msgs)


def test_instalar_modelo_sin_internet_no_rompe():
    import instalar_modelo
    msgs = []
    assert instalar_modelo.descargar("qwen2.5:7b-instruct", "http://127.0.0.1:9", msgs.append) is False
    assert any("Descargar modelo" in m for m in msgs)
