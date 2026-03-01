<#
.SYNOPSIS
    Compila, firma y empaqueta NotificadorBajasHitssApp como MSI.

.DESCRIPTION
    1. Publica la app como self-contained (win-x64)
    2. Crea un certificado autofirmado de code signing (si no existe)
    3. Firma el .exe con signtool
    4. Harvesta los archivos publicados con WiX
    5. Construye el MSI
    6. Firma el MSI

.PARAMETER Version
    Versión del instalador (por defecto: 1.0.0)

.EXAMPLE
    .\build-installer.ps1 -Version "1.0.1"

.NOTES
    Requisitos:
      - .NET 8 SDK
      - WiX 4 CLI: dotnet tool install --global wix
                   wix extension add WixToolset.UI.wixext
      - signtool.exe (incluido en Windows SDK)
#>
param(
    [string]$Version = "1.0.0"
)

$ErrorActionPreference = "Stop"
$Root    = $PSScriptRoot
$Publish = Join-Path $Root "publish"
$Setup   = Join-Path $Root "Setup"
$Dist    = Join-Path $Root "dist"
$CertPfx = Join-Path $Setup "BotHitss_CodeSign.pfx"
$CertPass = "BotHitss@Hitss2024"

# ─────────────────────────────────────────────────────────────────────────────
function Write-Step($msg) { Write-Host "`n> $msg" -ForegroundColor Cyan }
function Write-Ok($msg)   { Write-Host "  [OK] $msg" -ForegroundColor Green }
function Write-Warn($msg) { Write-Host "  [!] $msg" -ForegroundColor Yellow }

# ── 1. Publicar ──────────────────────────────────────────────────────────────
Write-Step "Publicando aplicación (self-contained win-x64)..."
if (Test-Path $Publish) { Remove-Item $Publish -Recurse -Force }

dotnet publish "$Root\NotificadorBajasHitssApp.csproj" `
    -c Release `
    -r win-x64 `
    --self-contained true `
    -p:PublishSingleFile=false `
    -p:Version=$Version `
    -o "$Publish"

Write-Ok "Publicado en: $Publish"

# ── 2. Certificado autofirmado ────────────────────────────────────────────────
Write-Step "Certificado de code signing..."

if (-not (Test-Path $CertPfx)) {
    Write-Host "  Generando certificado autofirmado (válido 5 años)..." -ForegroundColor Gray

    $cert = New-SelfSignedCertificate `
        -Type CodeSigning `
        -Subject "CN=Hitss Notificador Bajas, O=Hitss, C=MX" `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -HashAlgorithm SHA256 `
        -KeyUsage DigitalSignature `
        -NotAfter (Get-Date).AddYears(5)

    $pwd = ConvertTo-SecureString $CertPass -AsPlainText -Force
    Export-PfxCertificate -Cert $cert -FilePath $CertPfx -Password $pwd | Out-Null

    Write-Ok "Certificado guardado en: $CertPfx"
    Write-Host ""
    Write-Host "  NOTA: Este es un certificado AUTOFIRMADO." -ForegroundColor Yellow
    Write-Host "  Windows mostrará 'Editor desconocido' al instalar el MSI." -ForegroundColor Yellow
    Write-Host "  Para distribución pública considera un cert de DigiCert/Sectigo." -ForegroundColor Yellow
} else {
    Write-Ok "Usando certificado existente: $CertPfx"
}

# ── 3. Buscar signtool.exe ────────────────────────────────────────────────────
Write-Step "Buscando signtool.exe..."
$signtool = Get-Command signtool.exe -ErrorAction SilentlyContinue

if (-not $signtool) {
    # Intentar encontrarlo en Windows SDK
    $sdkPaths = @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\bin\*\x64\signtool.exe",
        "${env:ProgramFiles}\Windows Kits\10\bin\*\x64\signtool.exe"
    )
    foreach ($pattern in $sdkPaths) {
        $found = Get-Item $pattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($found) { $signtool = $found; break }
    }
}

if ($signtool) {
    $stSrc = if ($signtool.Source) { $signtool.Source } else { $signtool.FullName }
    Write-Ok "signtool encontrado: $stSrc"
} else {
    Write-Warn "signtool.exe no encontrado. Se omitirá la firma."
    Write-Warn "Instala Windows SDK para habilitar la firma: https://developer.microsoft.com/windows/downloads/windows-sdk/"
}

# ── 4. Firmar el ejecutable ───────────────────────────────────────────────────
$ExePath = Join-Path $Publish "NotificadorBajasHitss.exe"

if ($signtool -and (Test-Path $ExePath)) {
    Write-Step "Firmando ejecutable..."
    $stPath = if ($signtool -is [string]) { $signtool } elseif ($signtool.Source) { $signtool.Source } else { $signtool.FullName }
    & "$stPath" sign /fd SHA256 /f "$CertPfx" /p "$CertPass" /t http://timestamp.digicert.com "$ExePath"
    Write-Ok "Ejecutable firmado."
}

# ── 5. Instalar WiX 4 si no está (requiere WiX 4 para "wix harvest") ─────────
Write-Step "Verificando WiX 4..."
$wixCmd = Get-Command wix -ErrorAction SilentlyContinue
if (-not $wixCmd) {
    Write-Host "  Instalando WiX 4 CLI (dotnet tool install -g wix --version 4.0.0)..." -ForegroundColor Gray
    dotnet tool install --global wix --version 4.0.0
    wix extension add WixToolset.UI.wixext --global
    wix extension add WixToolset.Util.wixext --global
    $wixCmd = Get-Command wix
} else {
    wix extension add WixToolset.Util.wixext --global 2>$null
}
Write-Ok "WiX disponible."

# ── 6. Generar PublishedFiles.wxs (WiX 4 no tiene "harvest" en CLI) ────────────
Write-Step "Generando lista de archivos para WiX..."
$HarvestWxs = Join-Path $Setup "PublishedFiles.wxs"
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
[void]$sb.AppendLine('<Wix xmlns="http://wixtoolset.org/schemas/v4/wxs">')
[void]$sb.AppendLine('  <Fragment>')
[void]$sb.AppendLine('    <ComponentGroup Id="PublishFiles" Directory="INSTALLFOLDER">')
$files = Get-ChildItem -Path $Publish -Recurse -File | Where-Object { $_.FullName -notlike "*\ref\*" }
foreach ($f in $files) {
    $rel = $f.FullName.Substring($Publish.Length + 1).Replace("\", ".")
    $id = ($rel -replace '[^a-zA-Z0-9_.]', '_').Substring(0, [Math]::Min(72, $rel.Length))
    if ($id -match '^[0-9]') { $id = "f_$id" }
    $src = $f.FullName.Substring($Publish.Length + 1)
    $guid = [guid]::NewGuid().ToString("B").ToUpperInvariant()
    [void]$sb.AppendLine("      <Component Id=`"$id`" Guid=`"$guid`">")
    [void]$sb.AppendLine("        <File Id=`"$id`" Source=`"$src`" KeyPath=`"yes`" />")
    [void]$sb.AppendLine("      </Component>")
}
[void]$sb.AppendLine('    </ComponentGroup>')
[void]$sb.AppendLine('  </Fragment>')
[void]$sb.AppendLine('</Wix>')
$sb.ToString() | Set-Content -Path $HarvestWxs -Encoding UTF8
Write-Ok "Generado: $HarvestWxs ($($files.Count) archivos)"

# ── 7. Construir MSI ──────────────────────────────────────────────────────────
Write-Step "Construyendo MSI..."
New-Item -ItemType Directory -Force -Path $Dist | Out-Null
$MsiPath = Join-Path $Dist "NotificadorBajasHitss-$Version.msi"

wix build "$Setup\Package.wxs" "$HarvestWxs" `
    -arch x64 `
    -d "Version=$Version" `
    -b $Publish `
    -o "$MsiPath"

Write-Ok "MSI generado: $MsiPath"

# ── 8. Firmar el MSI ─────────────────────────────────────────────────────────
if ($signtool -and (Test-Path $MsiPath)) {
    Write-Step "Firmando MSI..."
    $stPath = if ($signtool -is [string]) { $signtool } elseif ($signtool.Source) { $signtool.Source } else { $signtool.FullName }
    & "$stPath" sign /fd SHA256 /f "$CertPfx" /p "$CertPass" /t http://timestamp.digicert.com "$MsiPath"
    Write-Ok "MSI firmado."
}

# ── Resumen ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Instalador listo: $MsiPath" -ForegroundColor Green
Write-Host "  Version: $Version" -ForegroundColor Green
Write-Host "  Instalacion: C:\ProgramData\Hitss\NotificadorBajasHitss (sin UAC al ejecutar)" -ForegroundColor Green
Write-Host "  Al iniciar, la app comprueba actualizaciones en el repositorio" -ForegroundColor Green
Write-Host "  Firmado con certificado autofirmado" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Para publicar en GitHub como release:" -ForegroundColor Gray
Write-Host ("    git tag v" + $Version + " ; git push origin v" + $Version) -ForegroundColor Gray
Write-Host "  Luego sube el MSI en: github.com/zipdatape/bothitss/releases" -ForegroundColor Gray
Write-Host ""
