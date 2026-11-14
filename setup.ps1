# Instala "Plumia" como complemento de Word, Excel y PowerPoint (catálogo de confianza).
# Uso: powershell -ExecutionPolicy Bypass -File setup.ps1
$ErrorActionPreference = "Stop"
$src     = $PSScriptRoot
$dest    = Join-Path $env:USERPROFILE "DeepSeekWordAddin"
$catalog = Join-Path $dest "catalog"
$certDir = Join-Path $dest "certs"
$pass    = "deepseek-word"
$catalogId = "{8e2f6a14-3c7d-4b59-a0e1-6d9c4f2b7a35}"

Write-Host "1/5 Copiando archivos a $dest"
New-Item -ItemType Directory -Force $dest, $catalog, $certDir, (Join-Path $dest "public\assets") | Out-Null
Copy-Item "$src\server.js" $dest -Force
Copy-Item "$src\public\*" (Join-Path $dest "public") -Recurse -Force
Copy-Item "$src\manifest.xml" $catalog -Force
Copy-Item "$src\manifest-outlook.xml" $dest -Force
# Archivos donde el usuario pega sus claves (se crean vacíos si no existen).
foreach ($p in "nvidia", "openrouter") {
  $kf = Join-Path $dest "clave-$p.txt"
  if (-not (Test-Path $kf)) { New-Item -ItemType File $kf | Out-Null }
}

Write-Host "2/5 Generando iconos"
Add-Type -AssemblyName System.Drawing
foreach ($s in 16, 32, 80) {
  $bmp = New-Object System.Drawing.Bitmap $s, $s
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = "AntiAlias"; $g.TextRenderingHint = "AntiAliasGridFit"
  $g.Clear([System.Drawing.Color]::FromArgb(77, 107, 254))
  $font = New-Object System.Drawing.Font("Segoe UI", [single]($s * 0.5), [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
  $sf = New-Object System.Drawing.StringFormat; $sf.Alignment = "Center"; $sf.LineAlignment = "Center"
  $g.DrawString("DS", $font, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF 0, 0, $s, $s), $sf)
  $bmp.Save((Join-Path $dest "public\assets\icon-$s.png"), [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose(); $bmp.Dispose()
}

Write-Host "3/5 Certificado HTTPS de localhost (Word exige HTTPS)"
$pfx = Join-Path $certDir "localhost.pfx"
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.FriendlyName -eq "Plumia localhost" } | Select-Object -First 1
if (-not $cert) {
  $cert = New-SelfSignedCertificate -DnsName "localhost" -CertStoreLocation Cert:\CurrentUser\My `
    -FriendlyName "Plumia localhost" -NotAfter (Get-Date).AddYears(5) -KeyAlgorithm RSA -KeyLength 2048 `
    -KeyExportPolicy Exportable -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.1")
}
Export-PfxCertificate -Cert $cert -FilePath $pfx -Password (ConvertTo-SecureString $pass -AsPlainText -Force) | Out-Null
$trusted = Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
if (-not $trusted) {
  Write-Host "   Windows pedirá confirmar que confías en este certificado: pulsa Sí."
  $cer = Join-Path $certDir "localhost.cer"
  Export-Certificate -Cert $cert -FilePath $cer | Out-Null
  Import-Certificate -FilePath $cer -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
}

Write-Host "4/5 Registrando el catálogo de complementos en Word"
$share = $null
try {
  if (-not (Get-SmbShare -Name "DeepSeekWordCatalog" -ErrorAction SilentlyContinue)) {
    New-SmbShare -Name "DeepSeekWordCatalog" -Path $catalog -ReadAccess "$env:USERDOMAIN\$env:USERNAME" | Out-Null
  }
  $share = "\\$env:COMPUTERNAME\DeepSeekWordCatalog"
} catch {
  $drive = $catalog.Substring(0, 1)
  $share = "\\$env:COMPUTERNAME\$drive`$\" + $catalog.Substring(3)
  Write-Host "   (sin permisos para compartir; se usa $share)"
}
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"
New-Item -Path $key -Force | Out-Null
New-ItemProperty -Path $key -Name Id    -Value $catalogId -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name Url   -Value $share     -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name Flags -Value 1          -PropertyType DWord  -Force | Out-Null
Write-Host "   Catálogo: $share"

Write-Host "5/5 Arranque automático del servidor local al iniciar sesión"
$vbs = Join-Path $dest "start-hidden.vbs"
@"
CreateObject("WScript.Shell").Run "cmd /c cd /d ""$dest"" && node server.js", 0, False
"@ | Set-Content -Path $vbs -Encoding ASCII
Copy-Item $vbs (Join-Path ([Environment]::GetFolderPath("Startup")) "DeepSeekWordAddin.vbs") -Force

Write-Host "Listo. Instalado en $dest"
"INSTALL_DIR=$dest"
