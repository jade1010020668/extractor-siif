<#
  Construye dist\ActaPrivada (app + Python + Ollama sin librerías NVIDIA) y el
  instalador dist\ActaPrivada-Setup.exe. Requiere Windows x64 con Python 3.11
  en el PATH (el mismo minor que el Python embebido). Inno Setup se instala
  con Chocolatey si no está.

  Uso:  pwsh packaging\windows\build.ps1 -Version 0.2.0
#>
param(
  [string]$Version = "0.0.0-dev",
  [string]$PythonVersion = "3.11.9",
  [string]$OllamaUrl = "https://github.com/ollama/ollama/releases/latest/download/ollama-windows-amd64.zip"
)
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$Root   = Resolve-Path (Join-Path $PSScriptRoot "..\..")      # acta-privada\
$Dist   = Join-Path $Root "dist"
$App    = Join-Path $Dist "ActaPrivada"
$Cache  = Join-Path $Dist "_descargas"
Remove-Item $App -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $App, $Cache | Out-Null

function Get-File($url, $dest) {
  if (-not (Test-Path $dest)) {
    Write-Host "Descargando $url"
    Invoke-WebRequest -Uri $url -OutFile $dest
  }
}

# 1. Código de la app (sin pruebas ni datos)
Write-Host "== Copiando la aplicación"
Copy-Item (Join-Path $Root "app.py"), (Join-Path $Root "launcher.py"), (Join-Path $Root "instalar_modelo.py"), (Join-Path $Root "README.md") $App
Copy-Item (Join-Path $Root "acta_privada") $App -Recurse
Get-ChildItem $App -Recurse -Directory -Filter "__pycache__" | Remove-Item -Recurse -Force
New-Item -ItemType Directory -Force (Join-Path $App "config"), (Join-Path $App ".streamlit") | Out-Null
Copy-Item (Join-Path $Root "config\*.example.json") (Join-Path $App "config")
Copy-Item (Join-Path $Root ".streamlit\config.toml") (Join-Path $App ".streamlit")
Copy-Item (Join-Path $PSScriptRoot "ActaPrivada.cmd") $App

# 2. Python embebido + dependencias
Write-Host "== Python $PythonVersion embebido"
$pyZip = Join-Path $Cache "python-$PythonVersion-embed-amd64.zip"
Get-File "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-embed-amd64.zip" $pyZip
$Py = Join-Path $App "python"
Expand-Archive $pyZip -DestinationPath $Py -Force
$tag = ($PythonVersion.Split(".")[0..1] -join "")
# sys.path: biblioteca estándar, site-packages y la carpeta de la app
Set-Content (Join-Path $Py "python$tag._pth") -Encoding ascii -Value @(
  "python$tag.zip", ".", "Lib\site-packages", "..", "import site")
$Site = Join-Path $Py "Lib\site-packages"
New-Item -ItemType Directory -Force $Site | Out-Null
python -m pip install --disable-pip-version-check --no-warn-script-location `
  --only-binary=:all: --target $Site -r (Join-Path $Root "requirements-app.txt")
if ($LASTEXITCODE) { throw "pip falló" }

# 3. Ollama sin librerías NVIDIA (CUDA ≈ 1,4 GB). Funciona con CPU y Vulkan.
Write-Host "== Ollama (sin CUDA)"
$olZip = Join-Path $Cache "ollama-windows-amd64.zip"
Get-File $OllamaUrl $olZip
$OlTmp = Join-Path $Cache "ollama"
Remove-Item $OlTmp -Recurse -Force -ErrorAction SilentlyContinue
Expand-Archive $olZip -DestinationPath $OlTmp -Force
Get-ChildItem (Join-Path $OlTmp "lib\ollama") -Directory |
  Where-Object { $_.Name -like "cuda*" -or $_.Name -like "rocm*" } |
  Remove-Item -Recurse -Force
Copy-Item $OlTmp (Join-Path $App "ollama") -Recurse
if (-not (Test-Path (Join-Path $App "ollama\ollama.exe"))) { throw "Falta ollama.exe en el paquete" }

# 4. Comprobación del paquete con el Python embebido
Write-Host "== Verificando el paquete"
& (Join-Path $Py "python.exe") -c "import acta_privada.structure, streamlit, docx, pandas; print('imports OK')"
if ($LASTEXITCODE) { throw "faltan dependencias en el paquete" }

$mb = [math]::Round(((Get-ChildItem $App -Recurse | Measure-Object Length -Sum).Sum / 1MB), 0)
Write-Host "Tamaño del paquete sin comprimir: $mb MB"

# 5. Instalador
Write-Host "== Instalador"
$iscc = @("${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe", "$env:ProgramFiles\Inno Setup 6\ISCC.exe") |
  Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) {
  choco install innosetup -y --no-progress | Out-Null
  $iscc = "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"
}
& $iscc "/DAppVersion=$Version" "/DDistDir=$App" "/DOutDir=$Dist" (Join-Path $PSScriptRoot "ActaPrivada.iss")
if ($LASTEXITCODE) { throw "Inno Setup falló" }
$exe = Join-Path $Dist "ActaPrivada-Setup.exe"
Write-Host ("Instalador: {0} ({1} MB)" -f $exe, [math]::Round((Get-Item $exe).Length / 1MB, 0))
