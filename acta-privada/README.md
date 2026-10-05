# Actas privadas — transcripción .docx → acta .docx (100 % local)

Sube el Word de la transcripción de la reunión y descarga el acta en Word con el
formato de la Sala Plena (datos generales, orden del día, desarrollo, firmas).

## ¿Ollama o reglas de negocio? Las dos, con roles distintos

| Tarea | Motor | Por qué |
|---|---|---|
| Participantes, cargos, fecha, horas de inicio/fin, orden del día, unanimidad, formato Word | **Reglas** (código determinista) | Exactitud y trazabilidad; no se le deja a un modelo decidir quién asistió o a qué hora terminó. |
| Redactar en tercera persona lo que cada Comisionado dijo (*«La Comisionada X, manifiesta que…»*), separar temas, resumir | **Ollama (IA local)** | Reescribir habla coloquial a lenguaje de acta no se puede hacer con reglas. |
| Sin IA instalada | **Solo reglas** (`--sin-ia`) | Borrador extractivo marcado `[REVISAR]`; sirve como respaldo si Ollama falla. |

**Ollama corre en tu propio equipo**: el texto nunca viaja a OpenAI, Anthropic, Google ni a ningún servicio en la nube.

## Cómo se protege la información

1. **Solo localhost.** `privacy.py` bloquea cualquier URL de Ollama que no sea `127.0.0.1`/`localhost`. Usar un servidor de tu red interna exige `ACTA_PERMITIR_RED_LOCAL=1` (decisión explícita).
2. **Sin proxies.** Las llamadas a Ollama ignoran `HTTP(S)_PROXY`, para que un proxy corporativo no vea el contenido.
3. **La interfaz solo escucha en `127.0.0.1`** (verificado) y sin telemetría (`gatherUsageStats=false`). No usa fuentes ni CDN externos.
4. **Nada se guarda en disco**: el archivo subido y el acta viven en memoria; «🗑️ Borrar todo» limpia la sesión. El CLI solo escribe el `.docx` de salida que tú indiques.
5. **El `.docx` generado no hereda metadatos** (autor, etc.) de la plantilla.
6. **Los documentos reales no entran a git**: `.gitignore` excluye `*.docx`, `config/roster.json`, `config/textos.json` y `plantilla/*`.
7. **Prueba automática**: `tests/test_acta.py::test_pipeline_completo_sin_salir_de_loopback` falla si el proceso intenta conectarse a cualquier IP que no sea loopback.

> ⚠️ **No despliegues esta app en Render, Streamlit Cloud ni similares** (el extractor SIIF de la raíz del repo sí está desplegado así; esta app no debe estarlo). Ejecútala solo en el computador de quien maneja la información, con disco cifrado (BitLocker/FileVault).

## Instalación (una vez)

```bash
# 1. Ollama: https://ollama.com/download   (luego, con internet, una sola vez:)
ollama pull qwen2.5:14b-instruct      # buen español; con poca RAM: qwen2.5:7b-instruct
# 2. App
cd acta-privada
pip install -r requirements.txt
cp config/roster.example.json config/roster.json     # editar con nombres y cargos reales
cp config/textos.example.json config/textos.json     # editar textos fijos de la entidad
# opcional: poner un acta anterior en plantilla/plantilla_acta.docx (logo, encabezado, pie)
./run.sh        # Windows: run.bat   → se abre en http://127.0.0.1:8501
```

Tras descargar el modelo puedes **desconectar el equipo de internet**: todo sigue funcionando (es la mejor forma de comprobar la privacidad).

Memoria orientativa: modelo 7B ≈ 8 GB de RAM; 14B ≈ 16 GB. Cambia el modelo en la barra lateral o con `ACTA_MODELO`.

## Uso

**Interfaz:** subir transcripción → revisar participantes → «Generar acta» → revisar/editar cada sección → descargar `.docx`.

**Línea de comandos:**
```bash
python -m acta_privada transcripcion.docx -o acta.docx --numero 074          # con IA local
python -m acta_privada transcripcion.docx -o acta.docx --numero 074 --sin-ia # solo reglas
```

## Formato de transcripción esperado

Exportación de Teams (o `.txt` equivalente): cada intervención como `Nombre   0:15` y el texto debajo. La primera línea (título con `… del 10 de septiembre de 2026, a las 1200 p.m.`) y la duración se usan para fecha y horas.

## Límites que debes conocer

- **El acta es un borrador asistido.** Siempre se revisa antes de firmar. Los datos dudosos se marcan `[VERIFICAR]`/`[REVISAR]` y la interfaz los cuenta.
- La IA puede equivocarse (números de decretos, artículos, fechas dichas de forma ambigua). Verifica contra el audio o los anexos.
- Las secciones *Solicitud* y *Decisión* solo se llenan si fueron explícitas en la reunión.
- La aprobación «por unanimidad» del orden del día solo se escribe si **todos** los comisionados expresaron conformidad; si no, se avisa.
- El texto fijo (p. ej. el párrafo de participación virtual) sale de `config/textos.json`.
- Calidad de redacción con modelos reales: depende del modelo elegido; ajusta los prompts en `acta_privada/structure.py` según tus actas.

## Pruebas

```bash
python -m pytest -q    # datos 100 % ficticios + Ollama simulado
```
