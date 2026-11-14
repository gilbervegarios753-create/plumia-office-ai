# Instalador de Plumia para Word, Excel y PowerPoint en Windows (versión en la nube).
#
# Qué hace:
#   1. Descarga el manifiesto de Plumia a %USERPROFILE%\Plumia\catalogo.
#   2. Registra esa carpeta como "catálogo de complementos de confianza" de Office (solo para tu usuario).
# No instala programas, no necesita Node.js ni certificados, y no envía ningún dato.
#
# Uso:   clic derecho en el archivo -> "Ejecutar con PowerShell"
#        Para quitarlo:  powershell -ExecutionPolicy Bypass -File instalar-plumia.ps1 -Desinstalar
param(
  [switch]$Desinstalar,
  [string]$Origen = "https://plumia-office-ai.vercel.app/descargas/plumia.xml",
  [switch]$Prueba  # solo para pruebas: no toca el registro
)
$ErrorActionPreference = "Stop"
$CatalogId = "{6a3c9e2b-7d41-4f58-9b0e-1c8d2f7a5e36}"
# Carpeta en el perfil del usuario: AppData puede no verse por la ruta de red que exige Office.
$Carpeta = Join-Path $env:USERPROFILE "Plumia\catalogo"
$Clave = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogId"

function Pausa { if (-not $Prueba) { Write-Host ""; Read-Host "Pulsa Enter para cerrar" | Out-Null } }

if ($Desinstalar) {
  if (-not $Prueba -and (Test-Path $Clave)) { Remove-Item $Clave -Recurse -Force }
  if (Test-Path $Carpeta) { Remove-Item $Carpeta -Recurse -Force }
  $padre = Split-Path $Carpeta -Parent
  if ((Test-Path $padre) -and -not (Get-ChildItem $padre -Force)) { Remove-Item $padre -Force }
  Write-Host "Plumia se ha quitado del catálogo. Si lo habías agregado en Word, quítalo también en Complementos -> Mis complementos." -ForegroundColor Green
  Pausa; return
}

Write-Host ""
Write-Host "  Instalando Plumia - tu equipo de IA en Office" -ForegroundColor Cyan
Write-Host ""

# 1. Descargar el manifiesto
Write-Host "1/3 Descargando Plumia..."
New-Item -ItemType Directory -Force $Carpeta | Out-Null
$destino = Join-Path $Carpeta "plumia.xml"
try {
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  if ($Origen -match '^https://') { Invoke-WebRequest -Uri $Origen -OutFile $destino -UseBasicParsing }
  else { Copy-Item $Origen $destino -Force }
} catch {
  Write-Host "No se pudo descargar Plumia: $($_.Exception.Message)" -ForegroundColor Red
  Write-Host "Revisa tu conexión a internet y vuelve a intentarlo."
  Pausa; exit 1
}
try { [xml]$xml = Get-Content $destino -Raw -Encoding UTF8 } catch { $xml = $null }
if (-not $xml -or -not $xml.OfficeApp.Id) {
  Write-Host "El archivo descargado no es un manifiesto válido de Plumia." -ForegroundColor Red
  Pausa; exit 1
}

# 2. Ruta de red de la carpeta (Office solo acepta catálogos en carpetas de red)
Write-Host "2/3 Preparando la carpeta de complementos..."
$unidad = $Carpeta.Substring(0, 1)
$ruta = "\\$env:COMPUTERNAME\$unidad`$\" + $Carpeta.Substring(3)
if (-not (Test-Path (Join-Path $ruta "plumia.xml"))) {
  Write-Host "Aviso: Windows no permite leer la carpeta por la ruta de red ($ruta)." -ForegroundColor Yellow
  Write-Host "Si Plumia no aparece en Word, pide ayuda en https://plumia-office-ai.vercel.app/soporte.html"
}

# 3. Registrar el catálogo de confianza de Office
Write-Host "3/3 Registrando Plumia en Office..."
if (-not $Prueba) {
  New-Item -Path $Clave -Force | Out-Null
  New-ItemProperty -Path $Clave -Name Id -Value $CatalogId -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $Clave -Name Url -Value $ruta -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $Clave -Name Flags -Value 1 -PropertyType DWord -Force | Out-Null
}

Write-Host ""
Write-Host "¡Listo! Ahora:" -ForegroundColor Green
Write-Host "  1. Cierra Word, Excel y PowerPoint por completo y vuelve a abrir Word."
Write-Host "  2. Ve a Inicio -> Complementos -> Más complementos."
Write-Host "  3. Abre la pestaña CARPETA COMPARTIDA, elige Plumia y pulsa Agregar."
Write-Host "  4. En Plumia, abre Ajustes y pega tu clave gratis (NVIDIA u OpenRouter)."
Write-Host ""
Write-Host "Ayuda: https://plumia-office-ai.vercel.app/soporte.html"
"RUTA=$ruta"
Pausa
