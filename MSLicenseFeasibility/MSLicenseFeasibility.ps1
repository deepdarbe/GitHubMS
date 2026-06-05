<#
.SYNOPSIS
    Microsoft Lisans Fizibilite Araci - ana orkestrator.
    (Microsoft licensing feasibility tool - main orchestrator.)

.DESCRIPTION
    Bir Windows Active Directory domain ortamini agentless (ajansiz) tarar ve
    Microsoft lisans ihtiyacini netlestirmek icin asagidakileri raporlar:
      - Aktif sunucu sayisi ve isletim sistemleri
      - Atanmis fiziksel core sayilari (Windows Server / SQL core lisanslama)
      - SQL Server instance + edition'lari
      - RDS sunuculari ve lisans modu (Per-User / Per-Device)
      - CAL hesabi icin etkin kullanici ve PC/cihaz adetleri

    Cikti: HTML yonetici raporu + CSV dosyalari (output klasoru).

    Windows PowerShell 5.1 ve PowerShell 7 ile uyumludur.

.PARAMETER Demo
    Gercek bir domain olmadan, sentetik ornek veriyle rapor uretir (onizleme).

.PARAMETER ComputerListFile
    Her satirinda bir sunucu adi bulunan duz metin dosyasi. AD kesfi yerine
    bu liste taranir. (AD modulu varsa kullanici/PC sayilari yine cekilir.)

.PARAMETER SearchBase
    AD sorgularini belirli bir OU ile sinirlar (DN). Ornek:
    "OU=Servers,DC=contoso,DC=local"

.PARAMETER StaleDays
    Bu gun sayisindan daha eski (LastLogon) AD nesneleri "pasif" sayilir.
    Varsayilan 90.

.PARAMETER Credential
    Uzak sunuculara baglanmak icin alternatif kimlik bilgisi.

.PARAMETER OutputDir
    Raporlarin yazilacagi klasor. Varsayilan: <script>\output

.PARAMETER PerHostTimeoutSec
    Sunucu basina CIM islem zaman asimi (sn). Varsayilan 20.

.PARAMETER NoHtml
    HTML raporu uretme, yalnizca CSV.

.PARAMETER OpenReport
    Rapor olusunca varsayilan tarayicida ac.

.EXAMPLE
    .\MSLicenseFeasibility.ps1
    Domain'i AD uzerinden kesfedip tum sunuculari tarar.

.EXAMPLE
    .\MSLicenseFeasibility.ps1 -Demo -OpenReport
    Ornek veriyle rapor uretir ve acar (test/onizleme).

.EXAMPLE
    .\MSLicenseFeasibility.ps1 -ComputerListFile .\servers.txt -Credential (Get-Credential)
#>
[CmdletBinding()]
param(
    [switch]$Demo,
    [string]$ComputerListFile,
    [string]$SearchBase,
    [int]$StaleDays = 90,
    [pscredential]$Credential,
    [string]$OutputDir,
    [int]$PerHostTimeoutSec = 20,
    [switch]$NoHtml,
    [switch]$OpenReport
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
if (-not $root) { $root = Split-Path -Parent $MyInvocation.MyCommand.Path }

# --- Modulleri yukle (dot-source) ---
# Ic ice Join-Path: hem Windows PowerShell 5.1 hem cross-platform uyumlu.
$libDir = Join-Path $root 'lib'
. (Join-Path $libDir 'LicensingEngine.ps1')
. (Join-Path $libDir 'Collectors.ps1')
. (Join-Path $libDir 'ReportWriter.ps1')

if (-not $OutputDir) { $OutputDir = Join-Path $root 'output' }
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

function Write-Banner {
    Write-Host ''
    Write-Host '  ============================================================' -ForegroundColor DarkCyan
    Write-Host '   Microsoft Lisans Fizibilite Araci  (agentless)' -ForegroundColor Cyan
    Write-Host '   Windows Server / CAL / RDS / SQL lisans envanteri' -ForegroundColor Cyan
    Write-Host '  ============================================================' -ForegroundColor DarkCyan
    Write-Host ''
}

function Get-ServerNameList {
    param([PSObject]$AdInventory)
    $names = foreach ($s in @($AdInventory.ActiveServers)) {
        if ($s.DNSHostName) { $s.DNSHostName } else { $s.Name }
    }
    return @($names)
}

Write-Banner

# ----------------------------------------------------------------------
# 1) Envanteri oku (Demo / Liste / AD kesfi)
# ----------------------------------------------------------------------
$inventory = $null

if ($Demo) {
    Write-Host '[i] DEMO modu: sentetik ornek veri kullaniliyor.' -ForegroundColor Yellow
    $inventory = New-DemoInventory
}
else {
    # Domain sayilari (kullanici / PC) AD'den; sunucu listesi AD veya dosyadan
    $adInv = $null
    $serverNames = @()

    if ($ComputerListFile) {
        if (-not (Test-Path $ComputerListFile)) { throw "Liste dosyasi bulunamadi: $ComputerListFile" }
        $serverNames = @(Get-Content -Path $ComputerListFile | Where-Object { $_ -and -not $_.StartsWith('#') } | ForEach-Object { $_.Trim() })
        Write-Host "[i] Liste dosyasindan $($serverNames.Count) sunucu okundu." -ForegroundColor Green
        try {
            $adInv = Get-DomainInventory -StaleDays $StaleDays -SearchBase $SearchBase
        } catch {
            Write-Warning "AD sorgusu yapilamadi (kullanici/PC sayilari 0 olacak): $($_.Exception.Message)"
        }
    }
    else {
        Write-Host '[i] Active Directory kesfi yapiliyor...' -ForegroundColor Green
        $adInv = Get-DomainInventory -StaleDays $StaleDays -SearchBase $SearchBase
        $serverNames = Get-ServerNameList -AdInventory $adInv
        Write-Host "[i] AD'de $($adInv.ServerCount) sunucu, $($adInv.EnabledUserCount) etkin kullanici, $($adInv.WorkstationCount) is istasyonu bulundu." -ForegroundColor Green
    }

    $dcNames = @()
    if ($adInv) { $dcNames = @($adInv.DomainControllers) }

    # 2) Her sunucudan envanter topla
    $collected = New-Object System.Collections.Generic.List[object]
    $idx = 0
    foreach ($name in $serverNames) {
        $idx++
        Write-Progress -Activity 'Sunucular taraniyor' -Status "$name ($idx/$($serverNames.Count))" -PercentComplete (($idx / [math]::Max($serverNames.Count,1)) * 100)
        $info = Get-ServerFullInventory -ComputerName $name -Credential $Credential -DomainControllers $dcNames -TimeoutSec $PerHostTimeoutSec
        $collected.Add($info)
        $stat = if ($info.Reachable -and -not $info.CollectionError) { 'OK' } else { 'ERISILEMEDI' }
        $col  = if ($stat -eq 'OK') { 'Gray' } else { 'DarkYellow' }
        Write-Host ("   [{0}] {1}  {2}" -f $stat, $name, ($(if ($info.OSCaption) { $info.OSCaption } else { $info.CollectionError }))) -ForegroundColor $col
    }
    Write-Progress -Activity 'Sunucular taraniyor' -Completed

    # Domain ozet nesnesi (Demo ile ayni sekil)
    $domain = [PSCustomObject]@{
        EnabledUserCount       = if ($adInv) { $adInv.EnabledUserCount } else { 0 }
        ServerCount            = $serverNames.Count
        ActiveServerCount      = @($collected | Where-Object { $_.Reachable }).Count
        WorkstationCount       = if ($adInv) { $adInv.WorkstationCount } else { 0 }
        ActiveWorkstationCount = if ($adInv) { $adInv.ActiveWorkstationCount } else { 0 }
        DomainControllers      = $dcNames
        StaleDays              = $StaleDays
    }

    $inventory = [PSCustomObject]@{
        Domain  = $domain
        Servers = $collected.ToArray()
    }
}

# ----------------------------------------------------------------------
# 3) Ozet + Rapor uret
# ----------------------------------------------------------------------
Write-Host ''
Write-Host '[i] Lisans hesabi yapiliyor...' -ForegroundColor Green
$summary = Get-FeasibilitySummary -Inventory $inventory
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

$csvFiles = Export-FeasibilityCsv -Summary $summary -Directory $OutputDir -Stamp $stamp
$htmlPath = $null
if (-not $NoHtml) {
    $htmlPath = Join-Path $OutputDir "LisansFizibilite_$stamp.html"
    New-HtmlReport -Summary $summary -Path $htmlPath | Out-Null
}

# ----------------------------------------------------------------------
# 4) Konsol ozeti
# ----------------------------------------------------------------------
Write-Host ''
Write-Host '  ---------------- LISANS IHTIYACI OZETI ----------------' -ForegroundColor Cyan
Write-Host ("   Toplam / erisilen sunucu      : {0} / {1}" -f $inventory.Domain.ServerCount, $summary.ReachableCount)
Write-Host ("   Fiziksel lisanslanabilir core : {0} core  ({1} x 2-core paket)" -f $summary.PhysicalCoreTotal, $summary.PhysicalPackTotal)
Write-Host ("   Sanal sunucu core (host bazli): {0} core" -f $summary.VirtualCoreTotal)
Write-Host ("   Windows Server CAL tavsiyesi  : {0} x {1}" -f $summary.Cal.RecommendedType, $summary.Cal.RecommendedCount)
Write-Host ("     (etkin kullanici: {0} / is istasyonu: {1})" -f $inventory.Domain.EnabledUserCount, $inventory.Domain.WorkstationCount)
if ($summary.Rds) {
    Write-Host ("   RDS CAL ({0})            : {1} x {2}" -f $summary.RdsMode, $summary.Rds.RecommendedType, $summary.Rds.RecommendedCount)
} else {
    Write-Host '   RDS                           : Session Host bulunamadi'
}
Write-Host ("   Ucretli SQL instance          : {0}" -f @($summary.SqlPaid).Count)
Write-Host '  -------------------------------------------------------' -ForegroundColor Cyan
Write-Host ''
Write-Host '[+] Olusturulan dosyalar:' -ForegroundColor Green
if ($htmlPath) { Write-Host "    HTML : $htmlPath" }
foreach ($f in $csvFiles) { Write-Host "    CSV  : $f" }
Write-Host ''
Write-Host '[!] UYARI: Bu bir tahmin aracidir; nihai lisans ihtiyacini yetkili' -ForegroundColor DarkYellow
Write-Host '    bir Microsoft lisans uzmaniyla teyit edin (SA, sozlesme, host->VM).' -ForegroundColor DarkYellow
Write-Host ''

if ($OpenReport -and $htmlPath) {
    try { Start-Process $htmlPath } catch { }
}
