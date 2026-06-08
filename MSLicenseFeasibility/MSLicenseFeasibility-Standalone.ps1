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

# ============================================================
#  GOMULU MODULLER  (OTOMATIK URETILDI - ELLE DUZENLEMEYIN)
#  Kaynak: lib/*.ps1   |   Uretici: build-standalone.ps1
# ============================================================

# -------- lib/LicensingEngine.ps1 --------
<#
.SYNOPSIS
    Microsoft lisans hesaplama motoru (Licensing calculation engine).

.DESCRIPTION
    Toplanan envanter verisini alir ve Microsoft'un resmi lisanslama
    kurallarina gore gerekli lisans adetlerini hesaplar:
      - Windows Server core lisansi (fiziksel sunucu + host-bazli VM)
      - Windows Server CAL (User / Device)
      - RDS CAL (Per-User / Per-Device)
      - SQL Server (Core modeli / Server+CAL modeli; fiziksel + VM)

    Kurallar Microsoft resmi lisanslama rehberlerinden (Windows Server 2022/2025
    ve SQL Server 2022 Licensing Guides, learn.microsoft.com) alinmistir.
    Bu arac yalnizca bir FIZIBILITE/tahmin aracidir; resmi lisans danismani
    veya Microsoft ile teyit edilmelidir.

    Bu dosya MSLicenseFeasibility.ps1 tarafindan dot-source edilir.
#>

# --------------------------------------------------------------------------
# Lisanslama sabitleri (Microsoft resmi esik degerleri)
# --------------------------------------------------------------------------
$Script:LicRules = [PSCustomObject]@{
    # Windows Server core-based lisanslama
    WS_MinCoresPerProcessor = 8     # Soket basina minimum core lisansi
    WS_MinCoresPerServer    = 16    # Sunucu basina minimum core lisansi
    WS_CorePackSize         = 2     # 2-core paketlerle satilir
    WS_StandardVmsPerSet    = 2     # Standard: lisans seti basina 2 OSE/VM
    WS_MinCoresPerVM        = 8     # Per-VM (SA ile) lisanslamada VM basina min

    # SQL Server core-based lisanslama
    SQL_MinCoresPerProcessor = 4    # Soket basina minimum core lisansi (fiziksel)
    SQL_CorePackSize         = 2    # 2-core paketlerle satilir
    SQL_MinCoresPerVM        = 4    # VM/OSE basina minimum (per-core)
    SQL_StandardEngineCoreCap = 24  # SQL Standard motoru: 4 soket veya 24 core'un kucugu

    # Standard -> Datacenter gecis esigi (VM yogunluguna gore tavsiye).
    DC_RecommendVmThreshold = 14
}

function Get-CoreLicenseCount {
    <#
    .SYNOPSIS
        Bir sunucu icin lisanslanmasi gereken core sayisini hesaplar.
    .DESCRIPTION
        Her soket icin "max(soketteki core, soket-minimum)" toplanir, sonra
        sunucu-minimumu uygulanir. Hyper-threading (mantiksal islemci) SAYILMAZ;
        yalnizca fiziksel core sayilir.
    #>
    param(
        [Parameter(Mandatory)][int]$Sockets,
        [Parameter(Mandatory)][int]$PhysicalCores,
        [int]$MinPerProcessor = 8,
        [int]$MinPerServer    = 16
    )

    if ($Sockets -lt 1)       { $Sockets = 1 }
    if ($PhysicalCores -lt 1) { $PhysicalCores = $Sockets * $MinPerProcessor }

    $coresPerSocket = [math]::Ceiling($PhysicalCores / $Sockets)
    $perSocketLicensed = $Sockets * [math]::Max($coresPerSocket, $MinPerProcessor)

    return [math]::Max($perSocketLicensed, $MinPerServer)
}

function Get-WindowsServerLicense {
    <#
    .SYNOPSIS
        Tek bir Windows sunucusu icin core lisans ihtiyacini hesaplar.
    .DESCRIPTION
        Fiziksel sunucu: 8/soket, 16/sunucu minimumu. VM (IsVirtual): gercek
        lisanslama HOST bazlidir (asagidaki Get-HostBasedWindowsLicense);
        per-VM (yalnizca SA/abonelik ile) basis = max(8, vCPU).
    #>
    param(
        [Parameter(Mandatory)][int]$Sockets,
        [Parameter(Mandatory)][int]$PhysicalCores,
        [int]$GuestVmCount = 0,
        [bool]$IsVirtualizationHost = $false,
        [bool]$IsVirtual = $false
    )

    $licensedCores = Get-CoreLicenseCount -Sockets $Sockets -PhysicalCores $PhysicalCores `
        -MinPerProcessor $Script:LicRules.WS_MinCoresPerProcessor `
        -MinPerServer    $Script:LicRules.WS_MinCoresPerServer

    $corePacks2 = [math]::Ceiling($licensedCores / $Script:LicRules.WS_CorePackSize)

    $standardSetsNeeded = 1
    $edition = 'Standard'
    $reason  = 'Fiziksel/tek sunucu - Standard yeterli (2 OSE hakki).'

    if ($IsVirtualizationHost -and $GuestVmCount -gt 0) {
        $standardSetsNeeded = [math]::Max(1, [math]::Ceiling($GuestVmCount / $Script:LicRules.WS_StandardVmsPerSet))
        if ($GuestVmCount -ge $Script:LicRules.DC_RecommendVmThreshold) {
            $edition = 'Datacenter'
            $reason  = "Yuksek VM yogunlugu ($GuestVmCount VM >= esik $($Script:LicRules.DC_RecommendVmThreshold)). Datacenter daha ekonomik ve sinirsiz VM saglar."
        } else {
            $reason  = "Sanallastirma host'u, $GuestVmCount VM. Standard ile $standardSetsNeeded lisans seti gerekir (her sette 2 VM)."
        }
    }

    # Per-VM (SA gerekli) lisans basis: VM basina min 8 core, vCPU kadar.
    $perVmCores = if ($IsVirtual) { [math]::Max($Script:LicRules.WS_MinCoresPerVM, $PhysicalCores) } else { $licensedCores }
    $licensedVia = if ($IsVirtual) { 'Host bazli (Datacenter/Standard) veya SA ile per-VM' } else { 'Fiziksel sunucu (core)' }

    return [PSCustomObject]@{
        IsVirtual             = $IsVirtual
        LicensedVia           = $licensedVia
        LicensedCores         = $licensedCores
        CorePacks_2Core       = [int]$corePacks2
        CorePacks_16Core      = [int][math]::Ceiling($licensedCores / 16)
        PerVmLicensableCores  = [int]$perVmCores       # SA ile per-VM senaryosu
        PerVmPacks_2Core      = [int][math]::Ceiling($perVmCores / 2)
        RecommendedEdition    = $edition
        StandardLicenseSets   = [int]$standardSetsNeeded
        StandardTotalCores    = [int]($licensedCores * $standardSetsNeeded)
        Reason                = $reason
    }
}

function Get-HostBasedWindowsLicense {
    <#
    .SYNOPSIS
        Fiziksel host(lar) icin host-bazli Windows Server lisansi (Datacenter / Standard).
    .DESCRIPTION
        Sanal sunucular fiziksel host'tan lisanslanir. Datacenter tum fiziksel
        core'lari bir kez lisanslar -> sinirsiz VM. Standard her lisans setiyle 2 VM
        (fazlasi icin core'lar yeniden lisanslanir). SA yoksa lisans tasinabilirligi
        yoktur; failover icin Standard her host'u tam kapsamak zorundadir, bu yuzden
        Datacenter cok daha ekonomik olur.
    #>
    param(
        [Parameter(Mandatory)][int[]]$HostPhysicalCores,  # her host'un fiziksel core sayisi
        [int]$GuestVmCount = 0,
        [bool]$NeedFailover = $true
    )

    $cores = @($HostPhysicalCores | ForEach-Object { [int]$_ })
    $hostCount = $cores.Count
    # Datacenter: her host'un tum fiziksel core'lari (host basina min 16).
    $dcPerHost = @($cores | ForEach-Object { [math]::Max($_, $Script:LicRules.WS_MinCoresPerServer) })
    $dcTotal = [int](($dcPerHost | Measure-Object -Sum).Sum)

    # Standard (pinned): VM'ler host'lara esit dagitilir, host basina ceil(vm/2) set.
    $stdPinned = 0
    if ($GuestVmCount -gt 0 -and $hostCount -gt 0) {
        $base = [math]::Floor($GuestVmCount / $hostCount)
        $rem  = $GuestVmCount % $hostCount
        for ($i = 0; $i -lt $hostCount; $i++) {
            $vmsOnHost = $base + $(if ($i -lt $rem) { 1 } else { 0 })
            if ($vmsOnHost -lt 1) { continue }
            $sets = [math]::Ceiling($vmsOnHost / $Script:LicRules.WS_StandardVmsPerSet)
            $stdPinned += $sets * $dcPerHost[$i]
        }
    }
    # Standard (failover): her host TUM VM'leri kapsamali (SA olmadan tasinma yok).
    $stdFailover = 0
    if ($GuestVmCount -gt 0) {
        $setsAll = [math]::Ceiling($GuestVmCount / $Script:LicRules.WS_StandardVmsPerSet)
        foreach ($hc in $dcPerHost) { $stdFailover += $setsAll * $hc }
    }

    $reason = if ($NeedFailover) {
        "Failover/HA senaryosu: Datacenter $dcTotal core ile $hostCount host'taki tum VM'leri + failover'i SINIRSIZ kapsar. Standard ayni failover icin ~$stdFailover core gerektirir (SA yoksa mobility yok)."
    } else {
        "Datacenter $dcTotal core ile tum VM'leri sinirsiz kapsar. Standard (pinned, failover yok) ~$stdPinned core - yalnizca VM'ler tek host'a sabitse mantiklidir."
    }

    return [PSCustomObject]@{
        HostCount               = $hostCount
        TotalPhysicalCores      = $dcTotal
        Datacenter_Cores        = $dcTotal
        Datacenter_Packs2       = [int][math]::Ceiling($dcTotal / 2)
        Datacenter_Packs16      = [int][math]::Ceiling($dcTotal / 16)
        Standard_Cores_Pinned   = [int]$stdPinned
        Standard_Cores_Failover = [int]$stdFailover
        NeedFailover            = $NeedFailover
        Recommended             = 'Datacenter'
        Reason                  = $reason
    }
}

function Get-SqlServerLicense {
    <#
    .SYNOPSIS
        Tek bir SQL Server instance'i icin lisans ihtiyacini hesaplar.
    .DESCRIPTION
        Iki model: Per-Core ve Server+CAL (Server+CAL yalnizca Standard, yeni anlasma).
        VM (IsVirtual) ise per-core TUM vCPU'lari lisanslar (OSE basina min 4);
        fiziksel ise soket basina min 4 uygulanir. SQL Standard 24 core ile sinirlidir.
    #>
    param(
        [Parameter(Mandatory)][int]$Sockets,
        [Parameter(Mandatory)][int]$PhysicalCores,
        [string]$Edition = 'Standard',
        [int]$KnownUserOrDeviceCount = 0,
        [bool]$IsVirtual = $false
    )

    if ($IsVirtual) {
        # VM: OSE'deki TUM sanal cekirdekler lisanslanir, OSE basina min 4 core.
        # Fiziksel soket-min kurali VM'de gecerli degildir.
        $licensedCores = [math]::Max($PhysicalCores, $Script:LicRules.SQL_MinCoresPerVM)
    } else {
        $licensedCores = Get-CoreLicenseCount -Sockets $Sockets -PhysicalCores $PhysicalCores `
            -MinPerProcessor $Script:LicRules.SQL_MinCoresPerProcessor `
            -MinPerServer    $Script:LicRules.SQL_MinCoresPerProcessor
    }

    # SQL Standard motoru en fazla 24 core kullanir; lisans da bununla sinirli.
    $capNote = ''
    if ($Edition -match 'Standard' -and $licensedCores -gt $Script:LicRules.SQL_StandardEngineCoreCap) {
        $capNote = " (Standard $($Script:LicRules.SQL_StandardEngineCoreCap) core ile sinirli)"
        $licensedCores = $Script:LicRules.SQL_StandardEngineCoreCap
    }

    $corePacks2 = [math]::Ceiling($licensedCores / $Script:LicRules.SQL_CorePackSize)

    $isEnterprise = $Edition -match 'Enterprise'
    $serverPlusCalAvailable = -not $isEnterprise

    $recommendedModel = 'Per-Core'
    $reason = 'Enterprise veya yuksek/bilinmeyen kullanici sayisi -> Per-Core model.'
    if ($serverPlusCalAvailable -and $KnownUserOrDeviceCount -gt 0) {
        # Kaba kiyas: Per-Core ~ corePacks2 paketi; Server+CAL ~ 1 sunucu + N CAL.
        # ~65 kullaniciya kadar Server+CAL genelde daha ekonomiktir.
        if ($KnownUserOrDeviceCount -le 65) {
            $recommendedModel = 'Server+CAL'
            $reason = "Standard ve $KnownUserOrDeviceCount kullanici (<=65) -> Server+CAL daha ekonomik (1 sunucu + $KnownUserOrDeviceCount SQL CAL)."
        } else {
            $reason = "Standard ancak $KnownUserOrDeviceCount kullanici (>65) -> Per-Core daha ekonomik ($corePacks2 x 2-core paketi, sinirsiz kullanici)."
        }
    }

    return [PSCustomObject]@{
        Edition                = $Edition
        IsVirtual              = $IsVirtual
        PerCore_LicensedCores  = $licensedCores
        PerCore_Packs_2Core    = [int]$corePacks2
        ServerPlusCalAvailable = $serverPlusCalAvailable
        ServerLicenses         = if ($serverPlusCalAvailable) { 1 } else { 0 }
        SqlCalsNeeded          = if ($serverPlusCalAvailable) { $KnownUserOrDeviceCount } else { 0 }
        RecommendedModel       = $recommendedModel
        Reason                 = ($reason + $capNote)
    }
}

function Get-CalRecommendation {
    <#
    .SYNOPSIS
        Windows Server CAL (User vs Device) tavsiyesi.
    #>
    param(
        [Parameter(Mandatory)][int]$UserCount,
        [Parameter(Mandatory)][int]$DeviceCount,
        [string]$MaxServerVersion = ''
    )

    $useUser = $UserCount -le $DeviceCount
    return [PSCustomObject]@{
        UserCount        = $UserCount
        DeviceCount      = $DeviceCount
        RecommendedType  = if ($useUser) { 'User CAL' } else { 'Device CAL' }
        RecommendedCount = if ($useUser) { $UserCount } else { $DeviceCount }
        UserCALs         = $UserCount
        DeviceCALs       = $DeviceCount
        VersionNote      = if ($MaxServerVersion) { "CAL surumu >= en yuksek sunucu surumu ($MaxServerVersion) olmalidir." } else { 'CAL surumu, erisilen en yuksek Windows Server surumune esit veya ustu olmalidir.' }
        Reason           = if ($useUser) { 'Kullanici sayisi <= cihaz sayisi: User CAL daha ekonomik.' } else { 'Cihaz sayisi < kullanici sayisi: Device CAL daha ekonomik (vardiyali/paylasimli cihazlar).' }
    }
}

function Get-RdsCalRecommendation {
    <#
    .SYNOPSIS
        RDS CAL ihtiyaci (Windows CAL'a EK olarak).
    #>
    param(
        [Parameter(Mandatory)][int]$RdsUserCount,
        [Parameter(Mandatory)][int]$RdsDeviceCount,
        [string]$DetectedMode = 'NotConfigured'
    )

    switch ($DetectedMode) {
        'PerUser'   { $type = 'RDS Per-User CAL';   $count = $RdsUserCount;   $reason = 'Sistemde Per-User modu tespit edildi.' }
        'PerDevice' { $type = 'RDS Per-Device CAL'; $count = $RdsDeviceCount; $reason = 'Sistemde Per-Device modu tespit edildi.' }
        default {
            if ($RdsUserCount -le $RdsDeviceCount) { $type = 'RDS Per-User CAL'; $count = $RdsUserCount; $reason = 'Mod yapilandirilmamis; kullanici sayisi az -> Per-User tavsiye.' }
            else { $type = 'RDS Per-Device CAL'; $count = $RdsDeviceCount; $reason = 'Mod yapilandirilmamis; cihaz sayisi az -> Per-Device tavsiye.' }
        }
    }

    return [PSCustomObject]@{
        DetectedMode     = $DetectedMode
        RecommendedType  = $type
        RecommendedCount = $count
        Note             = 'RDS CAL, Windows Server CAL''a EK olarak gereklidir (ikisi birden).'
        Reason           = $reason
    }
}

# Not: Bu dosya dot-source edildigi icin Export-ModuleMember kullanilmaz;
# tum fonksiyonlar cagiran kapsama (scope) otomatik aktarilir.

# -------- lib/Collectors.ps1 --------
<#
.SYNOPSIS
    Envanter toplama katmani (Inventory collection layer).

.DESCRIPTION
    Active Directory'den ve uzak sunuculardan (CIM/WMI uzerinden) lisans
    fizibilitesi icin gereken verileri toplar:
      - Domain bilgisayar/kullanici envanteri (sunucu / is istasyonu / DC ayrimi)
      - Isletim sistemi, fiziksel/sanal durum
      - Fiziksel soket ve core sayilari (lisanslanabilir core matematigi)
      - SQL Server instance + edition tespiti
      - RDS rolu, lisans modu ve kurulu RDS CAL paketleri

    Tum uzak cagrilar once erisilebilirlik testinden gecer ve try/catch ile
    sarmalanir; basarisiz hostlar ayri bir listede raporlanir.
    CIM (WS-MAN/WinRM) tercih edilir, basarisiz olursa DCOM'a geri donulur.

    Bu dosya MSLicenseFeasibility.ps1 tarafindan dot-source edilir.
#>

# --------------------------------------------------------------------------
# Yardimci: Guvenli CIM oturumu (once WSMan, sonra DCOM fallback)
# --------------------------------------------------------------------------
function New-SafeCimSession {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [pscredential]$Credential,
        [int]$TimeoutSec = 20
    )
    $common = @{ ComputerName = $ComputerName; ErrorAction = 'Stop' }
    if ($Credential) { $common['Credential'] = $Credential }

    # 1) WSMan (WinRM 5985) dene
    try {
        return New-CimSession @common -OperationTimeoutSec $TimeoutSec
    } catch {
        # 2) DCOM (RPC 135) fallback
        try {
            $dcom = New-CimSessionOption -Protocol Dcom
            return New-CimSession @common -SessionOption $dcom
        } catch {
            throw "CIM oturumu kurulamadi (WSMan ve DCOM basarisiz): $($_.Exception.Message)"
        }
    }
}

# --------------------------------------------------------------------------
# Yardimci: Host erisilebilir mi? (ping + yonetim portu fallback)
# --------------------------------------------------------------------------
function Test-HostReachable {
    param(
        [Parameter(Mandatory)][string]$ComputerName
    )
    # Hizli ICMP testi
    try {
        if (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction Stop) {
            return $true
        }
    } catch { }

    # ICMP kapali olabilir; yonetim portlarini dene
    foreach ($port in 5985, 135) {
        try {
            $r = Test-NetConnection -ComputerName $ComputerName -Port $port -WarningAction SilentlyContinue -ErrorAction Stop
            if ($r.TcpTestSucceeded) { return $true }
        } catch { }
    }
    return $false
}

# --------------------------------------------------------------------------
# AD: Domain geneli kullanici / bilgisayar envanteri
# --------------------------------------------------------------------------
function Get-DomainInventory {
    <#
    .SYNOPSIS
        Active Directory'den sunuculari, is istasyonlarini, DC'leri ve
        etkin kullanici sayisini toplar. Etkin olmayan/eski nesneleri eler.
    #>
    param(
        [int]$StaleDays = 90,
        [string]$SearchBase
    )

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw "ActiveDirectory modulu bulunamadi. RSAT-AD-PowerShell kurulu olmali veya -ComputerListFile kullanin."
    }
    Import-Module ActiveDirectory -ErrorAction Stop

    $baseParam = @{}
    if ($SearchBase) { $baseParam['SearchBase'] = $SearchBase }

    $staleDate = (Get-Date).AddDays(-$StaleDays)

    # Domain Controller'lar (otoriter)
    $dcNames = @()
    try { $dcNames = (Get-ADDomainController -Filter * -ErrorAction Stop).HostName } catch { }

    # Sunucular (etkin)
    $servers = Get-ADComputer -Filter "OperatingSystem -like '*Windows Server*' -and Enabled -eq 'True'" `
        -Properties OperatingSystem, OperatingSystemVersion, LastLogonDate, DNSHostName @baseParam

    # Is istasyonlari (etkin)
    $workstations = Get-ADComputer -Filter "OperatingSystem -notlike '*Server*' -and OperatingSystem -like '*Windows*' -and Enabled -eq 'True'" `
        -Properties OperatingSystem, OperatingSystemVersion, LastLogonDate @baseParam

    # Etkin kullanicilar (devre disi olmayanlar) - verimli LDAP bit filtresi
    $enabledUsers = @(Get-ADUser -LDAPFilter '(&(objectCategory=person)(objectClass=user)(!userAccountControl:1.2.840.113556.1.4.803:=2))' @baseParam)

    # Eski (stale) nesne ayrimi - CAL sayilarinin sismemesi icin
    $activeServers      = @($servers      | Where-Object { $_.LastLogonDate -ge $staleDate -or -not $_.LastLogonDate })
    $activeWorkstations = @($workstations | Where-Object { $_.LastLogonDate -ge $staleDate -or -not $_.LastLogonDate })

    return [PSCustomObject]@{
        Servers              = $servers
        ActiveServers        = $activeServers
        Workstations         = $workstations
        ActiveWorkstations   = $activeWorkstations
        DomainControllers    = $dcNames
        EnabledUserCount     = $enabledUsers.Count
        ServerCount          = @($servers).Count
        ActiveServerCount    = $activeServers.Count
        WorkstationCount     = @($workstations).Count
        ActiveWorkstationCount = $activeWorkstations.Count
        StaleDays            = $StaleDays
    }
}

# --------------------------------------------------------------------------
# OS + Core + Sanallastirma bilgisi (uzak)
# --------------------------------------------------------------------------
function Get-CoreAndOsInfo {
    param(
        [Parameter(Mandatory)][Microsoft.Management.Infrastructure.CimSession]$CimSession
    )
    $os  = Get-CimInstance -CimSession $CimSession -ClassName Win32_OperatingSystem -ErrorAction Stop
    $cs  = Get-CimInstance -CimSession $CimSession -ClassName Win32_ComputerSystem -ErrorAction Stop
    $cpu = @(Get-CimInstance -CimSession $CimSession -ClassName Win32_Processor -ErrorAction Stop)

    # Soket = Win32_Processor satir sayisi. Core = NumberOfCores toplami.
    $sockets = $cpu.Count
    $physicalCores = ($cpu | Measure-Object -Property NumberOfCores -Sum).Sum
    if (-not $physicalCores -or $physicalCores -lt 1) {
        # Cok eski OS: NumberOfCores yok -> soket sayisini core kabul et
        $physicalCores = $sockets
    }
    $logicalCores = ($cpu | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum

    # Sanal mi? Uretici/model imzasi + HypervisorPresent
    $vmSig = '{0} {1}' -f $cs.Manufacturer, $cs.Model
    $isVirtual = ($vmSig -match 'Virtual|VMware|KVM|QEMU|Xen|VirtualBox|Bochs|Amazon EC2|Google') `
        -or ($cs.Model -match 'Virtual Machine')

    return [PSCustomObject]@{
        OSCaption     = $os.Caption
        OSVersion     = $os.Version
        OSBuild       = $os.BuildNumber
        OSArchitecture= $os.OSArchitecture
        Manufacturer  = $cs.Manufacturer
        Model         = $cs.Model
        Sockets       = [int]$sockets
        PhysicalCores = [int]$physicalCores
        LogicalCores  = [int]$logicalCores
        HyperThreading= ($logicalCores -gt $physicalCores)
        IsVirtual     = [bool]$isVirtual
    }
}

# --------------------------------------------------------------------------
# SQL Server instance + edition tespiti (registry birincil, WMI provider yedek)
# --------------------------------------------------------------------------
function Get-SqlEditionClass {
    param([string]$Edition)
    # Ucretsiz / lisans gerektirmeyen edition'lar
    if ($Edition -match 'Express|Developer|Evaluation|Eval\b') { return 'Free' }
    if ($Edition -match 'Enterprise') { return 'Enterprise' }
    if ($Edition -match 'Standard|Business Intelligence|Web') { return 'Paid' }
    if ([string]::IsNullOrWhiteSpace($Edition)) { return 'Unknown' }
    return 'Paid'
}

function Get-SqlInventory {
    param(
        [Parameter(Mandatory)][Microsoft.Management.Infrastructure.CimSession]$CimSession
    )
    $instances = New-Object System.Collections.Generic.List[object]

    # --- Yontem 1: Uzak registry (StdRegProv uzerinden CIM) ---
    $HKLM = 2147483650
    try {
        $instKey = 'SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL'
        $enum = Invoke-CimMethod -CimSession $CimSession -Namespace 'root\cimv2' -ClassName StdRegProv `
            -MethodName EnumValues -Arguments @{ hDefKey = [uint32]$HKLM; sSubKeyName = $instKey } -ErrorAction Stop
        if ($enum.ReturnValue -eq 0 -and $enum.sNames) {
            for ($i = 0; $i -lt $enum.sNames.Count; $i++) {
                $instName = $enum.sNames[$i]
                # Instance'in dahili anahtarini al (MSSQLnn.INSTANCE)
                $val = Invoke-CimMethod -CimSession $CimSession -Namespace 'root\cimv2' -ClassName StdRegProv `
                    -MethodName GetStringValue -Arguments @{ hDefKey = [uint32]$HKLM; sSubKeyName = $instKey; sValueName = $instName } -ErrorAction Stop
                $internal = $val.sValue
                $setupKey = "SOFTWARE\Microsoft\Microsoft SQL Server\$internal\Setup"
                $edition = (Invoke-CimMethod -CimSession $CimSession -Namespace 'root\cimv2' -ClassName StdRegProv -MethodName GetStringValue -Arguments @{ hDefKey=[uint32]$HKLM; sSubKeyName=$setupKey; sValueName='Edition' } -ErrorAction SilentlyContinue).sValue
                $version = (Invoke-CimMethod -CimSession $CimSession -Namespace 'root\cimv2' -ClassName StdRegProv -MethodName GetStringValue -Arguments @{ hDefKey=[uint32]$HKLM; sSubKeyName=$setupKey; sValueName='Version' } -ErrorAction SilentlyContinue).sValue
                $patch = (Invoke-CimMethod -CimSession $CimSession -Namespace 'root\cimv2' -ClassName StdRegProv -MethodName GetStringValue -Arguments @{ hDefKey=[uint32]$HKLM; sSubKeyName=$setupKey; sValueName='PatchLevel' } -ErrorAction SilentlyContinue).sValue

                $instances.Add([PSCustomObject]@{
                    InstanceName = $instName
                    Edition      = $edition
                    EditionClass = Get-SqlEditionClass -Edition $edition
                    Version      = $version
                    PatchLevel   = $patch
                    Source       = 'Registry'
                })
            }
        }
    } catch { }

    # --- Yontem 2: SQL WMI provider (registry bos donerse) ---
    if ($instances.Count -eq 0) {
        try {
            $ns = @(Get-CimInstance -CimSession $CimSession -Namespace 'root\Microsoft\SqlServer' -ClassName __NAMESPACE -ErrorAction Stop) |
                Where-Object { $_.Name -like 'ComputerManagement*' }
            foreach ($cm in $ns) {
                $full = "root\Microsoft\SqlServer\$($cm.Name)"
                $props = @(Get-CimInstance -CimSession $CimSession -Namespace $full -ClassName SqlServiceAdvancedProperty -ErrorAction SilentlyContinue) |
                    Where-Object { $_.ServiceName -like 'MSSQL*' -and $_.PropertyName -in 'SKUNAME','VERSION','INSTANCENAME' }
                $byInstance = $props | Group-Object ServiceName
                foreach ($g in $byInstance) {
                    $sku = ($g.Group | Where-Object PropertyName -eq 'SKUNAME').PropertyStrValue
                    $ver = ($g.Group | Where-Object PropertyName -eq 'VERSION').PropertyStrValue
                    $instances.Add([PSCustomObject]@{
                        InstanceName = $g.Name
                        Edition      = $sku
                        EditionClass = Get-SqlEditionClass -Edition $sku
                        Version      = $ver
                        PatchLevel   = $null
                        Source       = 'WMI'
                    })
                }
            }
        } catch { }
    }

    return $instances.ToArray()
}

# --------------------------------------------------------------------------
# RDS rolu + lisans modu + kurulu CAL paketleri
# --------------------------------------------------------------------------
function Get-RdsInventory {
    param(
        [Parameter(Mandatory)][Microsoft.Management.Infrastructure.CimSession]$CimSession
    )
    $isSessionHost   = $false
    $isLicenseServer = $false
    $mode            = 'NotConfigured'
    $calPacks        = @()

    # RD Session Host / RD Licensing rol tespiti (Win32 ile, RemoteDesktop modulu gerekmeden)
    # TerminalServices namespace varsa Session Host kuruludur.
    $tss = $null
    try {
        $tss = Get-CimInstance -CimSession $CimSession -Namespace 'root\cimv2\TerminalServices' -ClassName Win32_TerminalServiceSetting -ErrorAction Stop
        $isSessionHost = $true
    } catch { }

    # Lisans modunu belirle. ONCE GPO/Policies registry anahtari (RCM'i ezer).
    $HKLM = 2147483650
    $policyMode = $null
    try {
        $pol = Invoke-CimMethod -CimSession $CimSession -Namespace 'root\cimv2' -ClassName StdRegProv -MethodName GetDWORDValue `
            -Arguments @{ hDefKey=[uint32]$HKLM; sSubKeyName='SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'; sValueName='LicensingMode' } -ErrorAction SilentlyContinue
        if ($pol.ReturnValue -eq 0 -and $null -ne $pol.uValue) { $policyMode = [int]$pol.uValue }
    } catch { }

    $modeCode = $null
    if ($null -ne $policyMode) {
        $modeCode = $policyMode
    } elseif ($tss -and $null -ne $tss.LicensingType) {
        $modeCode = [int]$tss.LicensingType
    } else {
        # RCM registry yedek
        try {
            $rcm = Invoke-CimMethod -CimSession $CimSession -Namespace 'root\cimv2' -ClassName StdRegProv -MethodName GetDWORDValue `
                -Arguments @{ hDefKey=[uint32]$HKLM; sSubKeyName='SYSTEM\CurrentControlSet\Control\Terminal Server\RCM\Licensing Core'; sValueName='LicensingMode' } -ErrorAction SilentlyContinue
            if ($rcm.ReturnValue -eq 0 -and $null -ne $rcm.uValue) { $modeCode = [int]$rcm.uValue }
        } catch { }
    }

    switch ($modeCode) {
        2       { $mode = 'PerDevice' }
        4       { $mode = 'PerUser' }
        5       { $mode = 'NotConfigured' }
        default { if ($null -ne $modeCode) { $mode = "Code$modeCode" } }
    }

    # Kurulu RDS CAL paketleri (lisans sunucusunda)
    try {
        $packs = @(Get-CimInstance -CimSession $CimSession -ClassName Win32_TSLicenseKeyPack -ErrorAction Stop) |
            Where-Object { $_.ProductType -ne 3 }   # 3 = BuiltIn/grace -> haric tut
        if ($packs) {
            $isLicenseServer = $true
            $calPacks = $packs | ForEach-Object {
                [PSCustomObject]@{
                    TypeAndModel     = $_.TypeAndModel
                    ProductVersion   = $_.ProductVersion
                    CalType          = switch ([int]$_.ProductType) { 0 {'Per Device'} 1 {'Per User'} default {"Type$($_.ProductType)"} }
                    TotalLicenses    = $_.TotalLicenses
                    IssuedLicenses   = $_.IssuedLicenses
                    AvailableLicenses= $_.AvailableLicenses
                }
            }
        }
    } catch { }

    return [PSCustomObject]@{
        IsRDSSessionHost   = $isSessionHost
        IsRDSLicenseServer = $isLicenseServer
        LicensingMode      = $mode
        InstalledCalPacks  = $calPacks
    }
}

# --------------------------------------------------------------------------
# Tek bir sunucudan tum envanteri topla
# --------------------------------------------------------------------------
function Get-ServerFullInventory {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [pscredential]$Credential,
        [string[]]$DomainControllers = @(),
        [int]$TimeoutSec = 20
    )

    $result = [PSCustomObject]@{
        ComputerName       = $ComputerName
        Reachable          = $false
        IsDomainController = ($DomainControllers -contains $ComputerName) -or ($DomainControllers | Where-Object { $_ -like "$ComputerName.*" }).Count -gt 0
        OSCaption          = $null
        OSVersion          = $null
        OSBuild            = $null
        Manufacturer       = $null
        Model              = $null
        IsVirtual          = $null
        Sockets            = 0
        PhysicalCores      = 0
        LogicalCores       = 0
        HyperThreading     = $false
        SQLInstances       = @()
        IsRDSSessionHost   = $false
        IsRDSLicenseServer = $false
        RDSLicensingMode   = 'N/A'
        RDSCalPacks        = @()
        CollectionError    = $null
    }

    if (-not (Test-HostReachable -ComputerName $ComputerName)) {
        $result.CollectionError = 'Erisilemiyor (ping/port basarisiz)'
        return $result
    }
    $result.Reachable = $true

    $session = $null
    try {
        $session = New-SafeCimSession -ComputerName $ComputerName -Credential $Credential -TimeoutSec $TimeoutSec

        $osInfo = Get-CoreAndOsInfo -CimSession $session
        $result.OSCaption     = $osInfo.OSCaption
        $result.OSVersion     = $osInfo.OSVersion
        $result.OSBuild       = $osInfo.OSBuild
        $result.Manufacturer  = $osInfo.Manufacturer
        $result.Model         = $osInfo.Model
        $result.IsVirtual     = $osInfo.IsVirtual
        $result.Sockets       = $osInfo.Sockets
        $result.PhysicalCores = $osInfo.PhysicalCores
        $result.LogicalCores  = $osInfo.LogicalCores
        $result.HyperThreading= $osInfo.HyperThreading

        $result.SQLInstances  = Get-SqlInventory -CimSession $session

        $rds = Get-RdsInventory -CimSession $session
        $result.IsRDSSessionHost   = $rds.IsRDSSessionHost
        $result.IsRDSLicenseServer = $rds.IsRDSLicenseServer
        $result.RDSLicensingMode   = $rds.LicensingMode
        $result.RDSCalPacks        = $rds.InstalledCalPacks
    }
    catch {
        $result.CollectionError = $_.Exception.Message
    }
    finally {
        if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
    }

    return $result
}

# --------------------------------------------------------------------------
# DEMO modu: gercek bir domain olmadan rapor onizlemesi icin sentetik envanter
# --------------------------------------------------------------------------
function New-DemoInventory {
    $servers = @(
        [PSCustomObject]@{ ComputerName='DC01'; Reachable=$true; IsDomainController=$true; OSCaption='Microsoft Windows Server 2022 Standard'; OSVersion='10.0.20348'; OSBuild='20348'; Manufacturer='VMware, Inc.'; Model='VMware Virtual Platform'; IsVirtual=$true; Sockets=2; PhysicalCores=8; LogicalCores=16; HyperThreading=$true; SQLInstances=@(); IsRDSSessionHost=$false; IsRDSLicenseServer=$false; RDSLicensingMode='N/A'; RDSCalPacks=@(); CollectionError=$null }
        [PSCustomObject]@{ ComputerName='DC02'; Reachable=$true; IsDomainController=$true; OSCaption='Microsoft Windows Server 2019 Standard'; OSVersion='10.0.17763'; OSBuild='17763'; Manufacturer='VMware, Inc.'; Model='VMware Virtual Platform'; IsVirtual=$true; Sockets=2; PhysicalCores=8; LogicalCores=16; HyperThreading=$true; SQLInstances=@(); IsRDSSessionHost=$false; IsRDSLicenseServer=$false; RDSLicensingMode='N/A'; RDSCalPacks=@(); CollectionError=$null }
        [PSCustomObject]@{ ComputerName='HV01'; Reachable=$true; IsDomainController=$false; OSCaption='Microsoft Windows Server 2022 Datacenter'; OSVersion='10.0.20348'; OSBuild='20348'; Manufacturer='Dell Inc.'; Model='PowerEdge R750'; IsVirtual=$false; Sockets=2; PhysicalCores=32; LogicalCores=64; HyperThreading=$true; SQLInstances=@(); IsRDSSessionHost=$false; IsRDSLicenseServer=$false; RDSLicensingMode='N/A'; RDSCalPacks=@(); CollectionError=$null }
        [PSCustomObject]@{ ComputerName='SQL01'; Reachable=$true; IsDomainController=$false; OSCaption='Microsoft Windows Server 2022 Standard'; OSVersion='10.0.20348'; OSBuild='20348'; Manufacturer='VMware, Inc.'; Model='VMware Virtual Platform'; IsVirtual=$true; Sockets=2; PhysicalCores=16; LogicalCores=32; HyperThreading=$true; SQLInstances=@( [PSCustomObject]@{ InstanceName='MSSQLSERVER'; Edition='Enterprise Edition'; EditionClass='Enterprise'; Version='15.0.2000.5'; PatchLevel='15.0.4345.5'; Source='Registry' } ); IsRDSSessionHost=$false; IsRDSLicenseServer=$false; RDSLicensingMode='N/A'; RDSCalPacks=@(); CollectionError=$null }
        [PSCustomObject]@{ ComputerName='SQL02'; Reachable=$true; IsDomainController=$false; OSCaption='Microsoft Windows Server 2019 Standard'; OSVersion='10.0.17763'; OSBuild='17763'; Manufacturer='VMware, Inc.'; Model='VMware Virtual Platform'; IsVirtual=$true; Sockets=1; PhysicalCores=8; LogicalCores=16; HyperThreading=$true; SQLInstances=@( [PSCustomObject]@{ InstanceName='MSSQL$APPDB'; Edition='Standard Edition'; EditionClass='Paid'; Version='15.0.2000.5'; PatchLevel='15.0.4345.5'; Source='Registry' } ); IsRDSSessionHost=$false; IsRDSLicenseServer=$false; RDSLicensingMode='N/A'; RDSCalPacks=@(); CollectionError=$null }
        [PSCustomObject]@{ ComputerName='RDS01'; Reachable=$true; IsDomainController=$false; OSCaption='Microsoft Windows Server 2022 Standard'; OSVersion='10.0.20348'; OSBuild='20348'; Manufacturer='VMware, Inc.'; Model='VMware Virtual Platform'; IsVirtual=$true; Sockets=2; PhysicalCores=12; LogicalCores=24; HyperThreading=$true; SQLInstances=@(); IsRDSSessionHost=$true; IsRDSLicenseServer=$true; RDSLicensingMode='PerUser'; RDSCalPacks=@( [PSCustomObject]@{ TypeAndModel='RDS Per User CAL'; ProductVersion='Windows Server 2022'; CalType='Per User'; TotalLicenses=100; IssuedLicenses=78; AvailableLicenses=22 } ); CollectionError=$null }
        [PSCustomObject]@{ ComputerName='APP01'; Reachable=$true; IsDomainController=$false; OSCaption='Microsoft Windows Server 2016 Standard'; OSVersion='10.0.14393'; OSBuild='14393'; Manufacturer='Dell Inc.'; Model='PowerEdge R640'; IsVirtual=$false; Sockets=1; PhysicalCores=4; LogicalCores=8; HyperThreading=$true; SQLInstances=@(); IsRDSSessionHost=$false; IsRDSLicenseServer=$false; RDSLicensingMode='N/A'; RDSCalPacks=@(); CollectionError=$null }
        [PSCustomObject]@{ ComputerName='FILE01'; Reachable=$true; IsDomainController=$false; OSCaption='Microsoft Windows Server 2022 Standard'; OSVersion='10.0.20348'; OSBuild='20348'; Manufacturer='VMware, Inc.'; Model='VMware Virtual Platform'; IsVirtual=$true; Sockets=2; PhysicalCores=10; LogicalCores=20; HyperThreading=$true; SQLInstances=@(); IsRDSSessionHost=$false; IsRDSLicenseServer=$false; RDSLicensingMode='N/A'; RDSCalPacks=@(); CollectionError=$null }
        [PSCustomObject]@{ ComputerName='OLD01'; Reachable=$false; IsDomainController=$false; OSCaption=$null; OSVersion=$null; OSBuild=$null; Manufacturer=$null; Model=$null; IsVirtual=$null; Sockets=0; PhysicalCores=0; LogicalCores=0; HyperThreading=$false; SQLInstances=@(); IsRDSSessionHost=$false; IsRDSLicenseServer=$false; RDSLicensingMode='N/A'; RDSCalPacks=@(); CollectionError='Erisilemiyor (ping/port basarisiz)' }
    )

    return [PSCustomObject]@{
        Domain = [PSCustomObject]@{
            EnabledUserCount       = 145
            ServerCount            = $servers.Count
            ActiveServerCount      = ($servers | Where-Object Reachable).Count
            WorkstationCount       = 160
            ActiveWorkstationCount = 152
            DomainControllers      = @('DC01.contoso.local','DC02.contoso.local')
            StaleDays              = 90
        }
        Servers = $servers
    }
}

# -------- lib/ReportWriter.ps1 --------
<#
.SYNOPSIS
    Rapor uretim katmani (HTML + CSV + Lisans Ihtiyac Matrisi).

.DESCRIPTION
    Toplanan envanteri lisans motorundan gecirir, ozetler/matris cikarir ve:
      - Kendine yeten (self-contained) bir HTML yonetici raporu
      - Ham veri ve ihtiyac matrisi icin CSV dosyalari
    uretir.

    Windows PowerShell 5.1 ile uyumludur (ternary / System.Web bagimliligi yok).
    Bu dosya MSLicenseFeasibility.ps1 tarafindan dot-source edilir ve
    LicensingEngine.ps1'in yuklenmis olmasini bekler.
#>

function ConvertTo-HtmlSafe {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    $s = [string]$Value
    $s = $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&#39;')
    return $s
}

function Get-FeasibilitySummary {
    <#
    .SYNOPSIS
        Envanteri lisans motorundan gecirip ozet/matris uretir.
    .PARAMETER PhysicalHostCores
        Fiziksel host(lar)in core sayilari (orn. @(24,24)). Verilirse host-bazli
        Windows (Datacenter/Standard) hesabi yapilir.
    #>
    param(
        [Parameter(Mandatory)][PSObject]$Inventory,
        [int[]]$PhysicalHostCores = @(),
        [bool]$NeedFailover = $true,
        [int]$RdsUserCountOverride = 0,
        [int]$RdsDeviceCountOverride = 0,
        [int]$SqlUserCountOverride = 0
    )

    $servers = @($Inventory.Servers)
    $reachable = @($servers | Where-Object { $_.Reachable -and -not $_.CollectionError })
    $winServers = @($reachable | Where-Object { $_.OSCaption -match 'Windows Server' })

    # Kullanici/cihaz sayilari (override > envanter)
    $sqlUsers   = if ($SqlUserCountOverride -gt 0)    { $SqlUserCountOverride }    else { [int]$Inventory.Domain.EnabledUserCount }
    $rdsUsers   = if ($RdsUserCountOverride -gt 0)    { $RdsUserCountOverride }    else { [int]$Inventory.Domain.EnabledUserCount }
    $rdsDevices = if ($RdsDeviceCountOverride -gt 0)  { $RdsDeviceCountOverride }  else { [int]$Inventory.Domain.WorkstationCount }

    # --- Windows Server core lisanslama (sunucu basina) ---
    $serverLicRows = foreach ($s in $winServers) {
        $lic = Get-WindowsServerLicense -Sockets ([int]$s.Sockets) -PhysicalCores ([int]$s.PhysicalCores) -IsVirtual:([bool]$s.IsVirtual)
        $edition = if ($s.OSCaption -match 'Datacenter') { 'Datacenter' } else { $lic.RecommendedEdition }
        [PSCustomObject]@{
            ComputerName       = $s.ComputerName
            OS                 = $s.OSCaption
            Type               = if ($s.IsVirtual) { 'Sanal (VM)' } else { 'Fiziksel' }
            Sockets            = $s.Sockets
            PhysicalCores      = $s.PhysicalCores
            LogicalCores       = $s.LogicalCores
            LicensableCores    = $lic.LicensedCores
            PerVmCores         = $lic.PerVmLicensableCores
            CorePacks_2Core    = $lic.CorePacks_2Core
            RecommendedEdition = $edition
            IsVirtual          = [bool]$s.IsVirtual
            IsDC               = [bool]$s.IsDomainController
        }
    }
    $serverLicRows = @($serverLicRows)

    $physical = @($serverLicRows | Where-Object { -not $_.IsVirtual })
    $virtual  = @($serverLicRows | Where-Object { $_.IsVirtual })

    $physCoreTotal = [int](($physical | Measure-Object -Property LicensableCores -Sum).Sum)
    $physPackTotal = [int](($physical | Measure-Object -Property CorePacks_2Core -Sum).Sum)
    $vmCoreTotal   = [int](($virtual  | Measure-Object -Property PhysicalCores -Sum).Sum)
    $vmPerVmCores  = [int](($virtual  | Measure-Object -Property PerVmCores -Sum).Sum)

    # --- Host-bazli Windows (fiziksel host core'u verilmisse) ---
    $hostBased = $null
    if (@($PhysicalHostCores).Count -gt 0) {
        $hostBased = Get-HostBasedWindowsLicense -HostPhysicalCores $PhysicalHostCores -GuestVmCount ($virtual.Count) -NeedFailover $NeedFailover
    }

    # --- SQL Server lisanslama ---
    $sqlRows = foreach ($s in $reachable) {
        foreach ($inst in @($s.SQLInstances)) {
            $lic = Get-SqlServerLicense -Sockets ([int]$s.Sockets) -PhysicalCores ([int]$s.PhysicalCores) `
                -Edition ([string]$inst.Edition) -KnownUserOrDeviceCount $sqlUsers -IsVirtual:([bool]$s.IsVirtual)
            [PSCustomObject]@{
                ComputerName     = $s.ComputerName
                Instance         = $inst.InstanceName
                Edition          = $inst.Edition
                EditionClass     = $inst.EditionClass
                Version          = $inst.Version
                PhysicalCores    = $s.PhysicalCores
                PerCore_Cores    = $lic.PerCore_LicensedCores
                PerCore_Packs2   = $lic.PerCore_Packs_2Core
                ServerPlusCAL    = $lic.ServerPlusCalAvailable
                RecommendedModel = $lic.RecommendedModel
                Licensable       = ($inst.EditionClass -ne 'Free')
            }
        }
    }
    $sqlRows = @($sqlRows)
    $sqlPaid = @($sqlRows | Where-Object { $_.Licensable })

    # Ayni OSE/VM'deki birden fazla instance -> tek lisans. Sunucu bazinda topla.
    $sqlByServer = @($sqlPaid | Group-Object ComputerName | ForEach-Object {
        $first = $_.Group[0]
        [PSCustomObject]@{
            ComputerName   = $_.Name
            Instances      = (($_.Group | ForEach-Object { $_.Instance }) -join ', ')
            InstanceCount  = $_.Count
            PerCore_Packs2 = $first.PerCore_Packs2
            ServerPlusCAL  = $first.ServerPlusCAL
        }
    })
    $sqlServerLicenses    = @($sqlByServer | Where-Object { $_.ServerPlusCAL }).Count
    $sqlPerCorePacksTotal = [int](($sqlByServer | Measure-Object -Property PerCore_Packs2 -Sum).Sum)

    # --- Windows Server CAL ---
    $cal = Get-CalRecommendation -UserCount ([int]$Inventory.Domain.EnabledUserCount) `
        -DeviceCount ([int]$Inventory.Domain.WorkstationCount)

    # --- RDS CAL ---
    $rdsHosts = @($reachable | Where-Object { $_.IsRDSSessionHost })
    $rdsLicServers = @($reachable | Where-Object { $_.IsRDSLicenseServer })
    $rdsMode = 'NotConfigured'
    $detected = @($rdsHosts | Where-Object { $_.RDSLicensingMode -in 'PerUser', 'PerDevice' } | Select-Object -First 1)
    if ($detected) { $rdsMode = $detected.RDSLicensingMode }
    $rds = $null
    if ($rdsHosts.Count -gt 0 -or $RdsUserCountOverride -gt 0) {
        $rds = Get-RdsCalRecommendation -RdsUserCount $rdsUsers -RdsDeviceCount $rdsDevices -DetectedMode $rdsMode
    }
    $installedCalPacks = foreach ($h in $rdsLicServers) { foreach ($p in @($h.RDSCalPacks)) { $p } }
    $installedCalPacks = @($installedCalPacks)

    # --- OS dagilimi ---
    $osDistribution = $winServers | Group-Object OSCaption | Sort-Object Count -Descending |
        ForEach-Object { [PSCustomObject]@{ OS = $_.Name; Count = $_.Count } }

    # --- LISANS IHTIYAC MATRISI ---
    $matrix = New-Object System.Collections.Generic.List[object]
    if ($hostBased) {
        $matrix.Add([PSCustomObject]@{ Item = 'Windows Server Datacenter (core)'; Qty = $hostBased.Datacenter_Cores; Unit = 'core'; Priority = 'Zorunlu'; Note = "$($hostBased.HostCount) fiziksel host, host-bazli. Tum VM + failover sinirsiz. = $($hostBased.Datacenter_Packs2)x 2-core paketi." })
    } else {
        $matrix.Add([PSCustomObject]@{ Item = 'Windows Server (fiziksel core)'; Qty = $physCoreTotal; Unit = 'core'; Priority = 'Zorunlu'; Note = "Fiziksel sunucular. VM'ler HOST bazli lisanslanir; -PhysicalHostsFile verirseniz Datacenter hesabi yapilir." })
    }
    $matrix.Add([PSCustomObject]@{ Item = "Windows Server $($cal.RecommendedType)"; Qty = $cal.RecommendedCount; Unit = 'CAL'; Priority = 'Zorunlu'; Note = $cal.Reason })
    if ($rds) {
        $matrix.Add([PSCustomObject]@{ Item = $rds.RecommendedType; Qty = $rds.RecommendedCount; Unit = 'CAL'; Priority = 'Zorunlu'; Note = "RD Session Host kullanicilari. Windows CAL'a EK." })
    }
    if ($sqlServerLicenses -gt 0) {
        $matrix.Add([PSCustomObject]@{ Item = 'SQL Server Standard - sunucu lisansi'; Qty = $sqlServerLicenses; Unit = 'sunucu lisansi'; Priority = 'Zorunlu'; Note = "OSE/VM basina 1 (tum instance dahil). Server+CAL modeli." })
        $matrix.Add([PSCustomObject]@{ Item = 'SQL Server CAL'; Qty = $sqlUsers; Unit = 'CAL'; Priority = 'Zorunlu'; Note = "SQL'e erisen kullanici/cihaz (multiplexing dahil). Windows CAL'dan ayri." })
        $matrix.Add([PSCustomObject]@{ Item = '(Alternatif) SQL Server Standard (core)'; Qty = $sqlPerCorePacksTotal; Unit = '2-core paketi'; Priority = 'Secenek'; Note = "Server+CAL yerine Per-Core: sinirsiz kullanici. ~65+ kullanici icin avantajli." })
    }

    $ret = [ordered]@{}
    $ret['Inventory']            = $Inventory
    $ret['ServerLicenseRows']    = $serverLicRows
    $ret['PhysicalServers']      = $physical
    $ret['VirtualServers']       = $virtual
    $ret['PhysicalCoreTotal']    = $physCoreTotal
    $ret['PhysicalPackTotal']    = $physPackTotal
    $ret['VirtualCoreTotal']     = $vmCoreTotal
    $ret['VirtualPerVmCoreTotal']= $vmPerVmCores
    $ret['HostBased']            = $hostBased
    $ret['SqlRows']              = $sqlRows
    $ret['SqlPaid']              = $sqlPaid
    $ret['SqlByServer']          = $sqlByServer
    $ret['SqlServerLicenses']    = $sqlServerLicenses
    $ret['SqlPerCorePacksTotal'] = $sqlPerCorePacksTotal
    $ret['SqlUsers']             = $sqlUsers
    $ret['Cal']                  = $cal
    $ret['Rds']                  = $rds
    $ret['RdsHosts']             = $rdsHosts
    $ret['RdsMode']              = $rdsMode
    $ret['RdsUsers']             = $rdsUsers
    $ret['InstalledRdsCalPacks'] = $installedCalPacks
    $ret['OsDistribution']       = @($osDistribution)
    $ret['ReachableCount']       = $reachable.Count
    $ret['UnreachableServers']   = @($servers | Where-Object { (-not $_.Reachable) -or $_.CollectionError })
    $ret['Matrix']               = $matrix.ToArray()
    return [pscustomobject]$ret
}

function New-HtmlReport {
    param(
        [Parameter(Mandatory)][PSObject]$Summary,
        [Parameter(Mandatory)][string]$Path
    )

    $inv = $Summary.Inventory
    $dom = $inv.Domain
    $now = Get-Date -Format 'yyyy-MM-dd HH:mm'

    $css = @'
<style>
  * { box-sizing:border-box; }
  body { font-family:Segoe UI,Arial,sans-serif; margin:0; background:#0b1220; color:#e5edff; }
  .wrap { max-width:1180px; margin:0 auto; padding:24px; }
  h1 { font-size:24px; margin:0 0 4px; }
  h2 { font-size:18px; margin:32px 0 12px; border-left:4px solid #2563eb; padding-left:10px; }
  .sub { color:#9fb0d0; font-size:13px; margin-bottom:18px; }
  .cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:14px; }
  .card { background:#16213a; border:1px solid #24304f; border-radius:10px; padding:16px; }
  .card .n { font-size:30px; font-weight:700; }
  .card .l { color:#9fb0d0; font-size:13px; margin-top:4px; }
  .card.accent .n { color:#60a5fa; } .card.good .n { color:#4ade80; } .card.warn .n { color:#fbbf24; }
  table { width:100%; border-collapse:collapse; margin-top:8px; font-size:13px; background:#16213a; border-radius:8px; overflow:hidden; }
  th,td { text-align:left; padding:9px 11px; border-bottom:1px solid #24304f; }
  th { background:#1c2949; color:#cdd9f5; font-weight:600; }
  tr:hover td { background:#1a2540; }
  .matrix th { background:#14532d; }
  .matrix td.qty { font-weight:700; color:#86efac; text-align:center; font-size:15px; }
  .pill { display:inline-block; padding:2px 8px; border-radius:999px; font-size:11px; font-weight:600; }
  .pill.phys { background:#1e3a8a; color:#bfdbfe; } .pill.vm { background:#374151; color:#d1d5db; }
  .pill.dc { background:#3730a3; color:#c7d2fe; }
  .pill.ent { background:#7c2d12; color:#fed7aa; } .pill.paid { background:#854d0e; color:#fde68a; } .pill.free { background:#14532d; color:#bbf7d0; }
  .pill.zor { background:#14532d; color:#bbf7d0; } .pill.sec { background:#854d0e; color:#fde68a; }
  .rec { background:#0c1f3a; border:1px solid #1e40af; border-radius:8px; padding:14px 16px; margin:10px 0; }
  .rec b { color:#93c5fd; }
  .disc { background:#3b1d1d; border:1px solid #7f1d1d; border-radius:8px; padding:12px 16px; color:#fecaca; font-size:12px; margin:16px 0; }
  .muted { color:#9fb0d0; }
  .foot { color:#9fb0d0; font-size:11px; margin-top:28px; border-top:1px solid #24304f; padding-top:12px; }
</style>
'@

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<!DOCTYPE html><html lang='tr'><head><meta charset='utf-8'><title>Microsoft Lisans Fizibilite Raporu</title>$css</head><body><div class='wrap'>")
    [void]$sb.Append("<h1>Microsoft Lisans Fizibilite Raporu</h1>")
    [void]$sb.Append("<div class='sub'>Olusturulma: $now &nbsp;|&nbsp; Domain DC: $(ConvertTo-HtmlSafe ($dom.DomainControllers -join ', ')) &nbsp;|&nbsp; Eski nesne esigi: $($dom.StaleDays) gun</div>")

    [void]$sb.Append("<div class='disc'><b>UYARI:</b> Bu rapor bir FIZIBILITE/tahmin aracidir. Lisanslama kurallari Microsoft resmi rehberlerine dayanir ancak nihai lisans ihtiyaci; Software Assurance, sozlesme tipi, sanallastirma haklari ve host&#8594;VM eslesmesi gibi etkenlere gore degisir. Satin alma oncesi yetkili bir Microsoft lisans uzmaniyla teyit edin.</div>")

    # --- LISANS IHTIYAC MATRISI ---
    if (@($Summary.Matrix).Count -gt 0) {
        [void]$sb.Append("<h2>Lisans Ihtiyac Matrisi</h2>")
        [void]$sb.Append("<table class='matrix'><tr><th>Urun / Lisans</th><th>Adet</th><th>Birim</th><th>Oncelik</th><th>Aciklama</th></tr>")
        foreach ($m in $Summary.Matrix) {
            $pc = if ($m.Priority -eq 'Zorunlu') { 'zor' } else { 'sec' }
            [void]$sb.Append("<tr><td><b>$(ConvertTo-HtmlSafe $m.Item)</b></td><td class='qty'>$($m.Qty)</td><td>$(ConvertTo-HtmlSafe $m.Unit)</td><td><span class='pill $pc'>$(ConvertTo-HtmlSafe $m.Priority)</span></td><td class='muted'>$(ConvertTo-HtmlSafe $m.Note)</td></tr>")
        }
        [void]$sb.Append("</table>")
    }

    # Executive cards
    [void]$sb.Append("<h2>Genel Bakis</h2><div class='cards'>")
    [void]$sb.Append("<div class='card accent'><div class='n'>$($dom.ServerCount)</div><div class='l'>Toplam Sunucu (AD)</div></div>")
    [void]$sb.Append("<div class='card'><div class='n'>$($Summary.ReachableCount)</div><div class='l'>Erisilen / Taranan Sunucu</div></div>")
    if ($Summary.HostBased) {
        [void]$sb.Append("<div class='card good'><div class='n'>$($Summary.HostBased.Datacenter_Cores)</div><div class='l'>Datacenter Core (host bazli)</div></div>")
    } else {
        [void]$sb.Append("<div class='card good'><div class='n'>$($Summary.PhysicalCoreTotal)</div><div class='l'>Fiziksel Lisanslanabilir Core</div></div>")
    }
    [void]$sb.Append("<div class='card'><div class='n'>$($dom.EnabledUserCount)</div><div class='l'>Etkin Kullanici (CAL)</div></div>")
    [void]$sb.Append("<div class='card warn'><div class='n'>$(@($Summary.SqlPaid).Count)</div><div class='l'>Ucretli SQL Instance</div></div>")
    [void]$sb.Append("</div>")

    # --- Host-based Windows ---
    if ($Summary.HostBased) {
        $hb = $Summary.HostBased
        [void]$sb.Append("<h2>Windows Server - Host Bazli Lisans</h2>")
        [void]$sb.Append("<div class='rec'><b>Tavsiye: Datacenter</b> = <b>$($hb.Datacenter_Cores)</b> core ($($hb.Datacenter_Packs2)x 2-core paketi), $($hb.HostCount) host. $(ConvertTo-HtmlSafe $hb.Reason)</div>")
        [void]$sb.Append("<table><tr><th>Senaryo</th><th>Toplam Core</th><th>Sonuc</th></tr>")
        [void]$sb.Append("<tr><td>Datacenter (onerilen)</td><td><b>$($hb.Datacenter_Cores)</b></td><td>Sinirsiz VM + failover</td></tr>")
        [void]$sb.Append("<tr><td>Standard (pinned, failover yok)</td><td>$($hb.Standard_Cores_Pinned)</td><td class='muted'>Sadece VM'ler tek host'a sabitse</td></tr>")
        [void]$sb.Append("<tr><td>Standard (failover ile)</td><td>$($hb.Standard_Cores_Failover)</td><td class='muted'>Onerilmez - Datacenter cok daha ucuz</td></tr>")
        [void]$sb.Append("</table>")
    }

    # --- Windows Server table ---
    [void]$sb.Append("<h2>Windows Server Envanteri &amp; Core</h2>")
    [void]$sb.Append("<table><tr><th>Sunucu</th><th>Isletim Sistemi</th><th>Tur</th><th>Soket</th><th>Fiziksel/vCPU</th><th>Per-VM Core</th><th>Edition</th></tr>")
    foreach ($r in $Summary.ServerLicenseRows) {
        $typePill = if ($r.IsVirtual) { "<span class='pill vm'>Sanal</span>" } else { "<span class='pill phys'>Fiziksel</span>" }
        $dcPill = if ($r.IsDC) { " <span class='pill dc'>DC</span>" } else { '' }
        $pvm = if ($r.IsVirtual) { "$($r.PerVmCores)" } else { '-' }
        [void]$sb.Append("<tr><td>$(ConvertTo-HtmlSafe $r.ComputerName)$dcPill</td><td>$(ConvertTo-HtmlSafe $r.OS)</td><td>$typePill</td><td>$($r.Sockets)</td><td>$($r.PhysicalCores)</td><td class='muted'>$pvm</td><td>$(ConvertTo-HtmlSafe $r.RecommendedEdition)</td></tr>")
    }
    [void]$sb.Append("</table>")
    [void]$sb.Append("<div class='sub' style='margin-top:8px'>Not: Sanal sunucularda 'Fiziksel/vCPU' atanmis vCPU'dur. Windows Server VM'leri <b>fiziksel host'tan</b> lisanslanir (Datacenter=sinirsiz VM, Standard=set basina 2 VM). 'Per-VM Core' yalnizca Software Assurance ile gecerli per-VM senaryosudur (VM basina min 8).</div>")

    # --- SQL table ---
    if (@($Summary.SqlRows).Count -gt 0) {
        [void]$sb.Append("<h2>SQL Server Envanteri &amp; Lisans</h2>")
        [void]$sb.Append("<table><tr><th>Sunucu</th><th>Instance</th><th>Edition</th><th>Tur</th><th>vCPU/Core</th><th>Per-Core (paket)</th><th>Model</th></tr>")
        foreach ($r in $Summary.SqlRows) {
            $cls = switch ($r.EditionClass) { 'Enterprise' { 'ent' } 'Free' { 'free' } default { 'paid' } }
            $clsLbl = switch ($r.EditionClass) { 'Free' { 'Ucretsiz' } 'Enterprise' { 'Enterprise' } default { 'Ucretli' } }
            [void]$sb.Append("<tr><td>$(ConvertTo-HtmlSafe $r.ComputerName)</td><td>$(ConvertTo-HtmlSafe $r.Instance)</td><td>$(ConvertTo-HtmlSafe $r.Edition) <span class='pill $cls'>$clsLbl</span></td><td class='muted'>$(if($r.IsVirtual){'VM'}else{'Fiziksel'})</td><td>$($r.PhysicalCores)</td><td><b>$($r.PerCore_Packs2)</b></td><td>$(ConvertTo-HtmlSafe $r.RecommendedModel)</td></tr>")
        }
        [void]$sb.Append("</table>")
        if (@($Summary.SqlByServer).Count -gt 0) {
            [void]$sb.Append("<div class='sub' style='margin-top:8px'>Sunucu bazinda (ayni VM'deki birden fazla instance TEK lisansla kapsanir): ")
            foreach ($g in $Summary.SqlByServer) {
                [void]$sb.Append("<b>$(ConvertTo-HtmlSafe $g.ComputerName)</b> ($($g.InstanceCount) instance) &rarr; Server+CAL: 1 sunucu + $($Summary.SqlUsers) CAL <i>veya</i> Per-Core: $($g.PerCore_Packs2) paket. &nbsp; ")
            }
            [void]$sb.Append("</div>")
        }
    }

    # --- RDS detail ---
    if ($Summary.RdsHosts.Count -gt 0) {
        [void]$sb.Append("<h2>RDS (Remote Desktop Services)</h2>")
        [void]$sb.Append("<table><tr><th>Sunucu</th><th>Session Host</th><th>Lisans Sunucusu</th><th>Lisans Modu</th><th>Kurulu CAL Paketleri</th></tr>")
        foreach ($h in $Summary.RdsHosts) {
            $packs = (@($h.RDSCalPacks) | ForEach-Object { "$($_.CalType): $($_.IssuedLicenses)/$($_.TotalLicenses)" }) -join '; '
            $sh = if ($h.IsRDSSessionHost) { 'Evet' } else { 'Hayir' }
            $ls = if ($h.IsRDSLicenseServer) { 'Evet' } else { 'Hayir' }
            [void]$sb.Append("<tr><td>$(ConvertTo-HtmlSafe $h.ComputerName)</td><td>$sh</td><td>$ls</td><td>$(ConvertTo-HtmlSafe $h.RDSLicensingMode)</td><td class='muted'>$(ConvertTo-HtmlSafe $packs)</td></tr>")
        }
        [void]$sb.Append("</table>")
    }

    # --- OS distribution ---
    [void]$sb.Append("<h2>Sunucu Isletim Sistemi Dagilimi</h2><table><tr><th>Isletim Sistemi</th><th>Adet</th></tr>")
    foreach ($o in $Summary.OsDistribution) {
        [void]$sb.Append("<tr><td>$(ConvertTo-HtmlSafe $o.OS)</td><td>$($o.Count)</td></tr>")
    }
    [void]$sb.Append("</table>")

    # --- Unreachable ---
    if (@($Summary.UnreachableServers).Count -gt 0) {
        [void]$sb.Append("<h2>Erisilemeyen / Hatali Sunucular</h2><table><tr><th>Sunucu</th><th>Durum</th></tr>")
        foreach ($u in $Summary.UnreachableServers) {
            $err = if ($u.CollectionError) { $u.CollectionError } else { 'Erisilemiyor' }
            [void]$sb.Append("<tr><td>$(ConvertTo-HtmlSafe $u.ComputerName)</td><td class='muted'>$(ConvertTo-HtmlSafe $err)</td></tr>")
        }
        [void]$sb.Append("</table>")
    }

    [void]$sb.Append("<div class='foot'>Kaynaklar (Microsoft resmi): Core-based licensing models; Windows Server 2022/2025 Licensing Guides; Client Access License (CAL); RDS CAL (learn.microsoft.com); SQL Server 2022 Licensing Guide. Bu arac yalnizca envanter ve tahmin amaclidir.</div>")
    [void]$sb.Append("</div></body></html>")

    $sb.ToString() | Out-File -FilePath $Path -Encoding UTF8
    return $Path
}

function Export-FeasibilityCsv {
    param(
        [Parameter(Mandatory)][PSObject]$Summary,
        [Parameter(Mandatory)][string]$Directory,
        [string]$Stamp
    )
    if (-not $Stamp) { $Stamp = Get-Date -Format 'yyyyMMdd_HHmmss' }
    $files = @()

    # Lisans ihtiyac matrisi (en onemli cikti)
    if (@($Summary.Matrix).Count -gt 0) {
        $pm = Join-Path $Directory "LisansMatrisi_$Stamp.csv"
        $Summary.Matrix | Select-Object @{n='Urun';e={$_.Item}}, @{n='Adet';e={$_.Qty}}, @{n='Birim';e={$_.Unit}}, @{n='Oncelik';e={$_.Priority}}, @{n='Aciklama';e={$_.Note}} |
            Export-Csv -Path $pm -NoTypeInformation -Encoding UTF8
        $files += $pm
    }

    $p1 = Join-Path $Directory "Servers_$Stamp.csv"
    $Summary.ServerLicenseRows | Select-Object ComputerName, OS, Type, Sockets, PhysicalCores, LogicalCores, LicensableCores, PerVmCores, CorePacks_2Core, RecommendedEdition, IsDC |
        Export-Csv -Path $p1 -NoTypeInformation -Encoding UTF8
    $files += $p1

    if (@($Summary.SqlRows).Count -gt 0) {
        $p2 = Join-Path $Directory "SQL_$Stamp.csv"
        $Summary.SqlRows | Select-Object ComputerName, Instance, Edition, EditionClass, Version, PhysicalCores, PerCore_Cores, PerCore_Packs2, ServerPlusCAL, RecommendedModel |
            Export-Csv -Path $p2 -NoTypeInformation -Encoding UTF8
        $files += $p2
    }

    $p3 = Join-Path $Directory "Summary_$Stamp.csv"
    $dcCore = if ($Summary.HostBased) { $Summary.HostBased.Datacenter_Cores } else { '' }
    [PSCustomObject]@{
        ToplamSunucu          = $Summary.Inventory.Domain.ServerCount
        ErisilenSunucu        = $Summary.ReachableCount
        Datacenter_Core_Host  = $dcCore
        FizikselLisansCore    = $Summary.PhysicalCoreTotal
        SanalvCPUToplam       = $Summary.VirtualCoreTotal
        EtkinKullanici        = $Summary.Inventory.Domain.EnabledUserCount
        CAL_Tavsiye           = "$($Summary.Cal.RecommendedType) x $($Summary.Cal.RecommendedCount)"
        RDS_CAL               = $(if ($Summary.Rds) { "$($Summary.Rds.RecommendedType) x $($Summary.Rds.RecommendedCount)" } else { 'Yok' })
        SQL_SunucuLisansi     = $Summary.SqlServerLicenses
        SQL_CAL               = $(if ($Summary.SqlServerLicenses -gt 0) { $Summary.SqlUsers } else { 0 })
        SQL_PerCorePaket_Alt  = $Summary.SqlPerCorePacksTotal
    } | Export-Csv -Path $p3 -NoTypeInformation -Encoding UTF8
    $files += $p3

    return $files
}

# ============================================================
#  ANA GOVDE
# ============================================================
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

