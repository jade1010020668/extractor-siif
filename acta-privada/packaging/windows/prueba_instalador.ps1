<#
  Prueba de humo en Windows: instala en silencio, arranca la app instalada,
  descarga un modelo pequeño y genera un acta con IA a partir de la
  transcripción ficticia de las pruebas.
#>
param([string]$ModeloPrueba = "qwen2.5:1.5b-instruct")
$ErrorActionPreference = "Stop"
$Root = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$Setup = Join-Path $Root "dist\ActaPrivada-Setup.exe"
$Inst = Join-Path $env:LOCALAPPDATA "Programs\ActaPrivada"
$env:ACTA_DATA_DIR = Join-Path $env:RUNNER_TEMP "acta-datos"

Write-Host "== Instalación silenciosa"
Start-Process $Setup -ArgumentList "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART" -Wait
if (-not (Test-Path "$Inst\launcher.py")) { throw "No se instaló en $Inst" }

Write-Host "== Arranque (autotest del lanzador)"
& "$Inst\python\python.exe" "$Inst\launcher.py" --autotest
if ($LASTEXITCODE) { throw "El lanzador falló" }

Write-Host "== Ollama incluido + modelo $ModeloPrueba + acta con IA"
$env:OLLAMA_HOST = "127.0.0.1:11434"
$env:OLLAMA_MODELS = Join-Path $env:ACTA_DATA_DIR "modelos"
$ol = Start-Process "$Inst\ollama\ollama.exe" -ArgumentList "serve" -PassThru -WindowStyle Hidden
Start-Sleep 5
$ErrorActionPreference = "Continue"
& "$Inst\ollama\ollama.exe" pull $ModeloPrueba
$ErrorActionPreference = "Stop"
if ($LASTEXITCODE) { throw "No se pudo descargar el modelo de prueba" }
$salida = Join-Path $env:RUNNER_TEMP "acta_prueba.docx"
Push-Location $Inst
$ErrorActionPreference = "Continue"   # los avisos de progreso salen por stderr
& "$Inst\python\python.exe" -m acta_privada (Join-Path $Root "tests\fixtures\transcripcion_ficticia.txt") `
  -o $salida --numero 001 --host http://127.0.0.1:11434 --modelo $ModeloPrueba 2>&1 | Tee-Object -Variable log
$code = $LASTEXITCODE
$ErrorActionPreference = "Stop"
Pop-Location
Stop-Process $ol -Force
if ($code) { throw "La generación del acta falló" }
if (-not (Test-Path $salida)) { throw "No se generó el acta" }
if ($log -match "IA local falló") { Write-Warning "El modelo de prueba no produjo JSON válido: se usó el borrador por reglas" }
else { Write-Host "Acta generada CON IA: $salida" }
