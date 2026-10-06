# CONTINUAR.md — guía para seguir desarrollando «Actas Privadas» en otro equipo

> Esta guía permite retomar el proyecto desde cero (en su computador, con Claude Code local u otro editor)
> sin depender de la conversación original. **No contiene datos reales de reuniones.**

## 1. Qué es y estado actual

App que recibe el **.docx de la transcripción de una reunión (Teams)** y devuelve el **acta en .docx**,
**sin que la información salga del equipo**.

- **Rama:** `claude/keen-lovelace-msm3jv` — PR en borrador: https://github.com/jade1010020668/extractor-siif/pull/1 (aún sin fusionar a `main`).
- **Carpeta:** `acta-privada/` (independiente del extractor SIIF de la raíz, que sí se despliega en Render/Streamlit Cloud; **esta app NO debe desplegarse en la nube**).
- **Versión publicada:** Release `acta-v0.1.0` → instalador `ActaPrivada-Setup.exe` (≈91 MB)
  https://github.com/jade1010020668/extractor-siif/releases/tag/acta-v0.1.0
- **Pruebas:** `python -m pytest -q` → 23 pasan. CI en Windows instala el `.exe`, lo arranca y genera un acta con un modelo real pequeño (`qwen2.5:1.5b`).
- **LO QUE NO ESTÁ VALIDADO:** la calidad de redacción con los modelos reales de uso (7B/14B) frente a un acta real. Es la prioridad nº 1 (ver §6).

## 2. Cómo ponerla a correr en su equipo (puerto local)

Requisitos: Python 3.11+, Git, [Ollama](https://ollama.com/download).

```bash
git clone -b claude/keen-lovelace-msm3jv https://github.com/jade1010020668/extractor-siif.git
cd extractor-siif/acta-privada
pip install -r requirements-app.txt pytest
ollama pull qwen2.5:7b-instruct          # 8 GB de RAM  (o qwen2.5:14b-instruct con 16 GB+)
python -m pytest -q                       # debe dar 23 passed
python launcher.py                        # abre http://127.0.0.1:8501 (solo local)
```

- Alternativas: `./run.sh` (Linux/Mac), `run.bat` (Windows) o `streamlit run app.py --server.address 127.0.0.1`.
- Sin Ollama la app funciona en modo **«Solo reglas»** (borrador extractivo `[REVISAR]`).
- CLI sin interfaz: `python -m acta_privada transcripcion.docx -o acta.docx --numero 074 --perfil 8gb` (o `--sin-ia`).
- Datos del usuario (config, plantilla, modelos): `%LOCALAPPDATA%\ActaPrivada` (Windows) o `~/.acta-privada`. Se cambia con `ACTA_DATA_DIR`.

## 3. Arquitectura (decisión clave: reglas + IA local)

| Tarea | Motor | Archivo |
|---|---|---|
| Leer transcripción, fecha, hora de inicio, duración | reglas | `acta_privada/transcript.py` |
| Reconocer personas y cargos (padrón con alias) | reglas | `roster.py` |
| Horas, aprobación del orden del día / unanimidad, avisos | reglas | `structure.py` |
| Detectar temas y redactar en 3.ª persona | **Ollama local** | `structure.py` + `llm.py` |
| Respaldo si la IA falla o no está | reglas (borrador extractivo) | `structure.py` |
| Generar el Word con el formato del acta | reglas | `docx_builder.py` |
| Perfiles por RAM (8 GB / 16 GB) | — | `perfiles.py` |
| Rutas de datos del usuario | — | `paths.py` |
| Bloqueos de privacidad | — | `privacy.py` |
| Arranque empaquetado (Ollama + UI) | — | `launcher.py` |
| Interfaz (Generar acta / Configuración) | — | `app.py` |

**Flujo de la IA** (`structure.py`): (A) detectar orden del día y temas con índice `[n]` de segmentos → (B1) por tema, redactar párrafos (`presentacion`/`discusion`/`cierre`) por tandas de `part_chars` → (B2) consolidar `pretension`/`solicitud`/`decision`. Todo con salida JSON por esquema. Las partes «duras» (quién asistió, horas, unanimidad) nunca las decide la IA.

**Formato del acta** (replicado por `docx_builder.py`): hoja Oficio 12240×18720 twips, Arial 11; título centrado (tipo de sesión / «ACTA No.» / fecha) + 5 tablas: datos generales y asistentes · orden del día · desarrollo (punto 1) · puntos 2…n con subtemas `n.k` (Pretensión, Presentación del tema, Discusión sobre el tema, Solicitud, Decisión) · firmas. Si el usuario sube una plantilla, se reutilizan solo encabezado/pie/estilos y se borran los metadatos.

## 4. Reglas de privacidad que NO se deben romper

1. Ollama solo en `127.0.0.1` (`privacy.validate_endpoint`); red interna solo con `ACTA_PERMITIR_RED_LOCAL=1`.
2. Modelos con «cloud» en el nombre están bloqueados (`privacy.validate_model`): se ejecutan en ollama.com.
3. Las llamadas a Ollama **ignoran proxies** (`llm._OPENER`).
4. UI solo en `127.0.0.1`, sin telemetría, sin CDN/fuentes externas. Nada de contenido de reuniones en disco ni en logs.
5. Nunca subir a git `.docx`, `config/roster.json`, `config/textos.json` ni `plantilla/*` (ya en `.gitignore`). Los tests usan **solo datos ficticios** (`tests/conftest.py`, `tests/fixtures/`).
6. La prueba `test_pipeline_completo_sin_salir_de_loopback` debe seguir pasando.

## 5. Empaquetado y publicación (Windows)

- Compilación local en Windows: `pwsh packaging/windows/build.ps1 -Version 0.2.0` → `dist/ActaPrivada-Setup.exe`
  (Python 3.11 embebido + dependencias fijadas en `requirements-app.txt` + Ollama **sin** librerías CUDA + Inno Setup).
- CI: `.github/workflows/acta-privada.yml` → pruebas (Linux) → build + prueba de instalación (Windows).
- **Publicar una versión:** subir una rama `acta-release/X.Y.Z` (p. ej. `git push origin HEAD:refs/heads/acta-release/0.2.0`) **o** un tag `acta-vX.Y.Z`. Crea la Release con el `.exe`.
  Desde la nube no se podían subir tags ni disparar el workflow a mano (403), por eso existe el disparo por rama.
- El instalador no está firmado: SmartScreen avisa («Más información → Ejecutar de todas formas»).

## 6. Pendientes, en orden de prioridad

1. **Medir la calidad con una acta real** (modelo 7B y 14B): generar con una transcripción real y comparar con el acta oficial. Ajustar los prompts de `structure.py` (`_SYS_TEMAS`, `_SYS_REDACCION`, `_SYS_CONSOLIDA`). Hasta ahora solo se probó con un modelo de 1.5B y datos ficticios.
2. Con IA, verificar que **Pretensión / Presentación / Discusión / Solicitud / Decisión** salgan bien separadas y con el estilo «La Comisionada X, manifiesta que…».
3. Transcripciones largas (> 1 h): comprobar tiempos y la tanda `part_chars`/`num_ctx` por perfil.
4. Mejoras posibles: anexos del acta, temas múltiples bajo varios puntos, exportar también PDF, firma digital del instalador, GPU NVIDIA (hoy el Ollama incluido va solo con CPU).
5. Fusionar el PR #1 a `main` cuando se valide.

## 7. Cosas que conviene saber (trampas encontradas)

- LibreOffice del entorno de nube no traía Writer: no se pudo renderizar el .docx a imagen; se validó releyéndolo con `python-docx`. En su equipo, abra el .docx en Word y compare visualmente con su acta de referencia.
- Windows reporta ~15,7 GB en un equipo de 16 GB: `perfiles.recomendado()` usa el umbral 15 GB.
- Exportación de Teams: cada intervención es `Nombre   0:15` seguido del texto; las líneas «ha iniciado/detenido la transcripción» se descartan.
- Los nombres en la transcripción vienen con errores (acentos, nombres cortados): el padrón con **alias** resuelve eso (`roster.match_person`).
- Los modelos tardan; con IA una acta de ~20 min de reunión toma varios minutos en CPU.

## 8. Cómo retomar con Claude Code en su equipo

1. Instale Claude Code y abra una terminal en `extractor-siif/acta-privada` (rama `claude/keen-lovelace-msm3jv`).
2. Primer mensaje sugerido:
   > «Lee `acta-privada/CONTINUAR.md` y `acta-privada/README.md`. Ejecuta `python -m pytest -q` y confirma que pasan. Vamos a trabajar el pendiente nº 1: tengo una transcripción real y el acta oficial; ayúdame a comparar y ajustar los prompts de `structure.py`. Recuerda las reglas de privacidad del §4: no subas ningún contenido real al repositorio.»
3. **Información sensible:** si usa Claude Code (o cualquier asistente en la nube) para ajustar prompts, **no le pegue transcripciones reales**. Use fragmentos ficticios o anonimizados, o trabaje con el modelo local y revise el resultado usted mismo.
