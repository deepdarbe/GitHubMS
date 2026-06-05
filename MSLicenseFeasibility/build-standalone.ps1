<#
.SYNOPSIS
    Tek-dosya (standalone) surumu uretir: MSLicenseFeasibility-Standalone.ps1

.DESCRIPTION
    lib/*.ps1 modullerini ve ana scriptin govdesini, kendine yeten tek bir
    .ps1 dosyasinda birlestirir. Uretilen dosya iex / irm / wget ile uzaktan
    (bellekten) calistirmaya uygundur cunku dot-source / $PSScriptRoot'a ihtiyac
    duymaz.

    Kaynak dosyalardan tek-yonlu uretim yapar; standalone dosyasini ELLE
    DUZENLEMEYIN, bunun yerine kaynaklari duzenleyip bu scripti tekrar calistirin.

.EXAMPLE
    .\build-standalone.ps1
#>
param([string]$OutFile)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
if (-not $root) { $root = Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $OutFile) { $OutFile = Join-Path $root 'MSLicenseFeasibility-Standalone.ps1' }

$mainPath = Join-Path $root 'MSLicenseFeasibility.ps1'
$main = Get-Content -Raw -Path $mainPath

$regionMarker = '#region MODULE-LOAD'
$bodyMarker   = '# === MAIN BODY ==='

if ($main.IndexOf($regionMarker) -lt 0 -or $main.IndexOf($bodyMarker) -lt 0) {
    throw "Isaretleyiciler bulunamadi ('$regionMarker' / '$bodyMarker'). Ana script degismis olabilir."
}

# Bas kisim: yardim + [CmdletBinding()] + param + ErrorActionPreference
$head = $main.Substring(0, $main.IndexOf($regionMarker)).TrimEnd()
# Govde: MAIN BODY isaretinden sonun sonuna kadar
$body = $main.Substring($main.IndexOf($bodyMarker)).TrimEnd()

$libDir = Join-Path $root 'lib'
$libs = 'LicensingEngine.ps1', 'Collectors.ps1', 'ReportWriter.ps1'

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine($head)
[void]$sb.AppendLine('')
[void]$sb.AppendLine('# ============================================================')
[void]$sb.AppendLine('#  GOMULU MODULLER  (OTOMATIK URETILDI - ELLE DUZENLEMEYIN)')
[void]$sb.AppendLine('#  Kaynak: lib/*.ps1   |   Uretici: build-standalone.ps1')
[void]$sb.AppendLine('# ============================================================')
foreach ($l in $libs) {
    $libText = (Get-Content -Raw -Path (Join-Path $libDir $l)).TrimEnd()
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("# -------- lib/$l --------")
    [void]$sb.AppendLine($libText)
}
[void]$sb.AppendLine('')
[void]$sb.AppendLine('# ============================================================')
[void]$sb.AppendLine('#  ANA GOVDE')
[void]$sb.AppendLine('# ============================================================')
[void]$sb.AppendLine($body)
[void]$sb.AppendLine('')

# CRLF satir sonu, BOM'suz UTF-8 (icerik saf ASCII)
$text = ($sb.ToString() -replace "`r`n", "`n") -replace "`n", "`r`n"
[System.IO.File]::WriteAllText($OutFile, $text, (New-Object System.Text.UTF8Encoding($false)))

$lines = ($text -split "`n").Count
Write-Host "[+] Uretildi: $OutFile  ($lines satir)" -ForegroundColor Green
