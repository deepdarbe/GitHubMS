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
    [switch]$OpenReport,
    # --- Host-bazli Windows lisanslama (sanal sunucular fiziksel host'tan lisanslanir) ---
    [string]$PhysicalHostsFile,   # CSV (PhysicalCores kolonu) ya da her satirda bir core sayisi
    [switch]$NoFailover,          # Varsayilan failover/HA varsayilir; pinned ise bunu verin
    # --- Bilinen sayilarla override (yoksa AD sayilari kullanilir) ---
    [int]$RdsUserCount = 0,
    [int]$RdsDeviceCount = 0,
    [int]$SqlUserCount = 0
)

$ErrorActionPreference = 'Stop'

#region MODULE-LOAD
# Bu blok yalnizca cok-dosyali kullanimda gereklidir. Standalone surum
# (build-standalone.ps1) bu bolgeyi kaldirip modulleri tek dosyaya gomer;
# boylece arac iex/irm/wget ile uzaktan (bellekten) calistirilabilir.
$__libRoot = $PSScriptRoot
if (-not $__libRoot) { $__libRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }
$__libDir = Join-Path $__libRoot 'lib'
. (Join-Path $__libDir 'LicensingEngine.ps1')
. (Join-Path $__libDir 'Collectors.ps1')
. (Join-Path $__libDir 'ReportWriter.ps1')
#endregion MODULE-LOAD

# === MAIN BODY ===
# Calisma kokunu coz: dosyadan calisirken script klasoru; iex/bellekten
# calisirken (PSScriptRoot bos) mevcut calisma dizini kullanilir.
$root = $PSScriptRoot
if (-not $root -and $MyInvocation.MyCommand.Path) { $root = Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $root) { $root = (Get-Location).Path }

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

# Fiziksel host core'larini oku (verilmisse) -> host-bazli Windows hesabi
$physicalHostCores = @()
if ($PhysicalHostsFile) {
    if (-not (Test-Path $PhysicalHostsFile)) { throw "Host dosyasi bulunamadi: $PhysicalHostsFile" }
    $rawHost = @(Get-Content -Path $PhysicalHostsFile | Where-Object { $_ -and -not $_.TrimStart().StartsWith('#') })
    $hasHeader = (($rawHost | Select-Object -First 1) -match 'PhysicalCores')
    if ($hasHeader) {
        $physicalHostCores = @(Import-Csv -Path $PhysicalHostsFile |
            Where-Object { "$($_.PhysicalCores)" -match '^\s*\d+\s*$' } |
            ForEach-Object { [int]$_.PhysicalCores })
    } else {
        $physicalHostCores = @($rawHost | ForEach-Object { [int]($_.Trim()) })
    }
    Write-Host ("[i] {0} fiziksel host okundu, toplam {1} core." -f $physicalHostCores.Count, (($physicalHostCores | Measure-Object -Sum).Sum)) -ForegroundColor Green
}

Write-Host ''
Write-Host '[i] Lisans hesabi yapiliyor...' -ForegroundColor Green
$summary = Get-FeasibilitySummary -Inventory $inventory `
    -PhysicalHostCores $physicalHostCores -NeedFailover (-not $NoFailover) `
    -RdsUserCountOverride $RdsUserCount -RdsDeviceCountOverride $RdsDeviceCount `
    -SqlUserCountOverride $SqlUserCount
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
Write-Host '  ---------------- LISANS IHTIYAC MATRISI ----------------' -ForegroundColor Cyan
foreach ($m in $summary.Matrix) {
    $col = if ($m.Priority -eq 'Zorunlu') { 'White' } else { 'DarkGray' }
    Write-Host ("   {0,-40} {1,5}  {2}" -f $m.Item, $m.Qty, $m.Unit) -ForegroundColor $col
}
Write-Host '  --------------------------------------------------------' -ForegroundColor Cyan
Write-Host ("   Toplam/erisilen sunucu: {0}/{1}    Ucretli SQL instance: {2}" -f $inventory.Domain.ServerCount, $summary.ReachableCount, @($summary.SqlPaid).Count) -ForegroundColor DarkGray
if ($summary.HostBased) {
    Write-Host ("   Windows host-bazli  -> Datacenter: {0} core ({1}x 2-core) | Standard(pinned): {2} core" -f $summary.HostBased.Datacenter_Cores, $summary.HostBased.Datacenter_Packs2, $summary.HostBased.Standard_Cores_Pinned) -ForegroundColor DarkGray
}
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
