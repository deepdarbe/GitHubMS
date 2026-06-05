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
      - Windows Server core lisansi (Standard / Datacenter)
      - Windows Server CAL (User / Device)
      - RDS CAL (Per-User / Per-Device)
      - SQL Server (Core modeli / Server+CAL modeli)

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

    # SQL Server core-based lisanslama
    SQL_MinCoresPerProcessor = 4    # Soket basina minimum core lisansi
    SQL_CorePackSize         = 2    # 2-core paketlerle satilir
    SQL_MinCoresPerVM        = 4    # VM basina minimum (per-core)
    SQL_StandardEngineCoreCap = 24  # SQL Standard motoru: 4 soket veya 24 core'un kucugu

    # Standard -> Datacenter gecis esigi (VM yogunluguna gore tavsiye).
    # Datacenter fiyati ~ Standard'in 8 kati oldugundan kirilma noktasi
    # genelde host basina ~12-14 VM civarindadir. Yapilandirilabilir.
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

    # Soketler arasi esit dagilim varsayilir (per-soket detay yoksa).
    $coresPerSocket = [math]::Ceiling($PhysicalCores / $Sockets)
    $perSocketLicensed = $Sockets * [math]::Max($coresPerSocket, $MinPerProcessor)

    return [math]::Max($perSocketLicensed, $MinPerServer)
}

function Get-WindowsServerLicense {
    <#
    .SYNOPSIS
        Tek bir Windows sunucusu icin core lisans ihtiyacini hesaplar.
    #>
    param(
        [Parameter(Mandatory)][int]$Sockets,
        [Parameter(Mandatory)][int]$PhysicalCores,
        [int]$GuestVmCount = 0,        # Bu host uzerinde calisan misafir VM sayisi (Hyper-V host ise)
        [bool]$IsVirtualizationHost = $false
    )

    $licensedCores = Get-CoreLicenseCount -Sockets $Sockets -PhysicalCores $PhysicalCores `
        -MinPerProcessor $Script:LicRules.WS_MinCoresPerProcessor `
        -MinPerServer    $Script:LicRules.WS_MinCoresPerServer

    $corePacks2 = [math]::Ceiling($licensedCores / $Script:LicRules.WS_CorePackSize)

    # Standard edition: 2 VM / lisans seti. Daha fazla VM icin tum core'lar
    # yeniden lisanslanir (stacking). Datacenter: sinirsiz VM.
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

    return [PSCustomObject]@{
        LicensedCores         = $licensedCores
        CorePacks_2Core       = [int]$corePacks2
        CorePacks_16Core      = [int][math]::Ceiling($licensedCores / 16)
        RecommendedEdition    = $edition
        StandardLicenseSets   = [int]$standardSetsNeeded
        # Standard secilirse satin alinacak toplam core lisansi (stacking dahil)
        StandardTotalCores    = [int]($licensedCores * $standardSetsNeeded)
        Reason                = $reason
    }
}

function Get-SqlServerLicense {
    <#
    .SYNOPSIS
        Tek bir SQL Server instance'i icin lisans ihtiyacini hesaplar.
    .DESCRIPTION
        Iki model degerlendirilir: Per-Core ve Server+CAL.
        Server+CAL yalnizca Standard edition icin (yeni anlasmalarda) gecerlidir.
    #>
    param(
        [Parameter(Mandatory)][int]$Sockets,
        [Parameter(Mandatory)][int]$PhysicalCores,
        [string]$Edition = 'Standard',
        [int]$KnownUserOrDeviceCount = 0   # SQL'e erisen bilinen kullanici/cihaz sayisi (Server+CAL icin)
    )

    $licensedCores = Get-CoreLicenseCount -Sockets $Sockets -PhysicalCores $PhysicalCores `
        -MinPerProcessor $Script:LicRules.SQL_MinCoresPerProcessor `
        -MinPerServer    $Script:LicRules.SQL_MinCoresPerProcessor   # SQL'de sabit sunucu-min yok; soket-min surer

    $corePacks2 = [math]::Ceiling($licensedCores / $Script:LicRules.SQL_CorePackSize)

    $isEnterprise = $Edition -match 'Enterprise'
    $serverPlusCalAvailable = -not $isEnterprise   # Server+CAL yalnizca Standard (yeni anlasma)

    # Model tavsiyesi: bilinen/kucuk kullanici sayisi -> Server+CAL; aksi halde Per-Core.
    $recommendedModel = 'Per-Core'
    $reason = 'Enterprise veya yuksek/bilinmeyen kullanici sayisi -> Per-Core model.'
    if ($serverPlusCalAvailable -and $KnownUserOrDeviceCount -gt 0) {
        # Kaba kiyas: Server+CAL ~ kucuk kullanici sayilarinda avantajli.
        $recommendedModel = 'Server+CAL'
        $reason = "Standard edition ve bilinen kullanici/cihaz sayisi ($KnownUserOrDeviceCount) -> Server+CAL degerlendirilmeli."
    }

    return [PSCustomObject]@{
        Edition                = $Edition
        PerCore_LicensedCores  = $licensedCores
        PerCore_Packs_2Core    = [int]$corePacks2
        ServerPlusCalAvailable = $serverPlusCalAvailable
        ServerLicenses         = if ($serverPlusCalAvailable) { 1 } else { 0 }
        SqlCalsNeeded          = if ($serverPlusCalAvailable) { $KnownUserOrDeviceCount } else { 0 }
        RecommendedModel       = $recommendedModel
        Reason                 = $reason
    }
}

function Get-CalRecommendation {
    <#
    .SYNOPSIS
        Windows Server CAL (User vs Device) tavsiyesi.
    .DESCRIPTION
        CAL ag genelinde gecerlidir (sunucu basina degil). Kullanici sayisi
        cihaz sayisindan az ise User CAL, aksi halde Device CAL daha ekonomiktir.
    #>
    param(
        [Parameter(Mandatory)][int]$UserCount,
        [Parameter(Mandatory)][int]$DeviceCount,
        [string]$MaxServerVersion = ''
    )

    $useUser = $UserCount -le $DeviceCount
    return [PSCustomObject]@{
        UserCount          = $UserCount
        DeviceCount        = $DeviceCount
        RecommendedType    = if ($useUser) { 'User CAL' } else { 'Device CAL' }
        RecommendedCount   = if ($useUser) { $UserCount } else { $DeviceCount }
        UserCALs           = $UserCount
        DeviceCALs         = $DeviceCount
        VersionNote        = if ($MaxServerVersion) { "CAL surumu >= en yuksek sunucu surumu ($MaxServerVersion) olmalidir." } else { 'CAL surumu, erisilen en yuksek Windows Server surumune esit veya ustu olmalidir.' }
        Reason             = if ($useUser) { 'Kullanici sayisi <= cihaz sayisi: User CAL daha ekonomik.' } else { 'Cihaz sayisi < kullanici sayisi: Device CAL daha ekonomik (vardiyali/paylasimli cihazlar).' }
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
        [string]$DetectedMode = 'NotConfigured'   # PerUser | PerDevice | NotConfigured
    )

    switch ($DetectedMode) {
        'PerUser'   { $type = 'RDS Per-User CAL';   $count = $RdsUserCount;   $reason = 'Sistemde Per-User modu tespit edildi.' }
        'PerDevice' { $type = 'RDS Per-Device CAL'; $count = $RdsDeviceCount; $reason = 'Sistemde Per-Device modu tespit edildi.' }
        default {
            # Mod yapilandirilmamis: kullanici<=cihaz ise Per-User oner.
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
    Rapor uretim katmani (HTML + CSV).

.DESCRIPTION
    Toplanan envanteri lisans motorundan gecirir, ozetler cikarir ve:
      - Kendine yeten (self-contained) bir HTML yonetici raporu
      - Ham veri icin CSV dosyalari
    uretir.

    Windows PowerShell 5.1 ile uyumludur (ternary operatoru / System.Web
    bagimliligi kullanilmaz). Bu dosya MSLicenseFeasibility.ps1 tarafindan
    dot-source edilir ve LicensingEngine.ps1'in yuklenmis olmasini bekler.
#>

# Basit, bagimsiz HTML encode (System.Web gerektirmez; 5.1 + 7 uyumlu)
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
        Envanteri lisans motorundan gecirip ozet/aggregate uretir.
    #>
    param(
        [Parameter(Mandatory)][PSObject]$Inventory
    )

    $servers = @($Inventory.Servers)
    $reachable = @($servers | Where-Object { $_.Reachable -and -not $_.CollectionError })
    $winServers = @($reachable | Where-Object { $_.OSCaption -match 'Windows Server' })

    # --- Windows Server core lisanslama (sunucu basina) ---
    $serverLicRows = foreach ($s in $winServers) {
        $lic = Get-WindowsServerLicense -Sockets ([int]$s.Sockets) -PhysicalCores ([int]$s.PhysicalCores)
        $edition = if ($s.OSCaption -match 'Datacenter') { 'Datacenter' } else { $lic.RecommendedEdition }
        [PSCustomObject]@{
            ComputerName       = $s.ComputerName
            OS                 = $s.OSCaption
            Type               = if ($s.IsVirtual) { 'Sanal (VM)' } else { 'Fiziksel' }
            Sockets            = $s.Sockets
            PhysicalCores      = $s.PhysicalCores
            LogicalCores       = $s.LogicalCores
            LicensableCores    = $lic.LicensedCores
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
    $vmCoreTotal   = [int](($virtual  | Measure-Object -Property LicensableCores -Sum).Sum)

    # --- SQL Server lisanslama ---
    $sqlRows = foreach ($s in $reachable) {
        foreach ($inst in @($s.SQLInstances)) {
            $lic = Get-SqlServerLicense -Sockets ([int]$s.Sockets) -PhysicalCores ([int]$s.PhysicalCores) `
                -Edition ([string]$inst.Edition) -KnownUserOrDeviceCount ([int]$Inventory.Domain.EnabledUserCount)
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
    if ($rdsHosts.Count -gt 0) {
        $rds = Get-RdsCalRecommendation -RdsUserCount ([int]$Inventory.Domain.EnabledUserCount) `
            -RdsDeviceCount ([int]$Inventory.Domain.WorkstationCount) -DetectedMode $rdsMode
    }
    $installedCalPacks = foreach ($h in $rdsLicServers) { foreach ($p in @($h.RDSCalPacks)) { $p } }
    $installedCalPacks = @($installedCalPacks)

    # --- OS dagilimi ---
    $osDistribution = $winServers | Group-Object OSCaption | Sort-Object Count -Descending |
        ForEach-Object { [PSCustomObject]@{ OS = $_.Name; Count = $_.Count } }

    return [PSCustomObject]@{
        Inventory            = $Inventory
        ServerLicenseRows    = $serverLicRows
        PhysicalServers      = $physical
        VirtualServers       = $virtual
        PhysicalCoreTotal    = $physCoreTotal
        PhysicalPackTotal    = $physPackTotal
        VirtualCoreTotal     = $vmCoreTotal
        SqlRows              = $sqlRows
        SqlPaid              = $sqlPaid
        Cal                  = $cal
        Rds                  = $rds
        RdsHosts             = $rdsHosts
        RdsMode              = $rdsMode
        InstalledRdsCalPacks = $installedCalPacks
        OsDistribution       = @($osDistribution)
        ReachableCount       = $reachable.Count
        UnreachableServers   = @($servers | Where-Object { -not $_.Reachable -or $_.CollectionError })
    }
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
  .pill { display:inline-block; padding:2px 8px; border-radius:999px; font-size:11px; font-weight:600; }
  .pill.phys { background:#1e3a8a; color:#bfdbfe; } .pill.vm { background:#374151; color:#d1d5db; }
  .pill.dc { background:#3730a3; color:#c7d2fe; }
  .pill.ent { background:#7c2d12; color:#fed7aa; } .pill.paid { background:#854d0e; color:#fde68a; } .pill.free { background:#14532d; color:#bbf7d0; }
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

    # Disclaimer
    [void]$sb.Append("<div class='disc'><b>UYARI:</b> Bu rapor bir FIZIBILITE/tahmin aracidir. Lisanslama kurallari Microsoft resmi rehberlerine dayanir ancak nihai lisans ihtiyaci; Software Assurance, sozlesme tipi, sanallastirma haklari ve host&#8594;VM eslesmesi gibi etkenlere gore degisir. Satin alma oncesi yetkili bir Microsoft lisans uzmaniyla teyit edin.</div>")

    # Executive cards
    [void]$sb.Append("<div class='cards'>")
    [void]$sb.Append("<div class='card accent'><div class='n'>$($dom.ServerCount)</div><div class='l'>Toplam Sunucu (AD)</div></div>")
    [void]$sb.Append("<div class='card'><div class='n'>$($Summary.ReachableCount)</div><div class='l'>Erisilen / Taranan Sunucu</div></div>")
    [void]$sb.Append("<div class='card good'><div class='n'>$($Summary.PhysicalCoreTotal)</div><div class='l'>Fiziksel Lisanslanabilir Core (WS)</div></div>")
    [void]$sb.Append("<div class='card'><div class='n'>$($dom.EnabledUserCount)</div><div class='l'>Etkin Kullanici (CAL)</div></div>")
    [void]$sb.Append("<div class='card'><div class='n'>$($dom.WorkstationCount)</div><div class='l'>Is Istasyonu / PC (CAL)</div></div>")
    [void]$sb.Append("<div class='card warn'><div class='n'>$(@($Summary.SqlPaid).Count)</div><div class='l'>Ucretli SQL Instance</div></div>")
    [void]$sb.Append("</div>")

    # --- Recommendations summary ---
    [void]$sb.Append("<h2>Lisans Ihtiyaci Ozeti (Tahmini)</h2>")
    [void]$sb.Append("<div class='rec'><b>Windows Server (Core):</b> Fiziksel sunucularda toplam <b>$($Summary.PhysicalCoreTotal)</b> lisanslanabilir core (= <b>$($Summary.PhysicalPackTotal)</b> adet 2-core paketi). Min. kurallar: soket basina 8, sunucu basina 16 core. Sanal sunuculardaki toplam $($Summary.VirtualCoreTotal) core, host bazli lisanslanir (asagidaki nota bakin).</div>")
    [void]$sb.Append("<div class='rec'><b>Windows Server CAL:</b> Tavsiye = <b>$(ConvertTo-HtmlSafe $Summary.Cal.RecommendedType)</b> &times; <b>$($Summary.Cal.RecommendedCount)</b>. ($($Summary.Cal.UserCALs) kullanici / $($Summary.Cal.DeviceCALs) cihaz). $(ConvertTo-HtmlSafe $Summary.Cal.Reason) <span class='muted'>$(ConvertTo-HtmlSafe $Summary.Cal.VersionNote)</span></div>")
    if ($Summary.Rds) {
        $packInfo = ''
        if (@($Summary.InstalledRdsCalPacks).Count -gt 0) {
            $tot = [int]((@($Summary.InstalledRdsCalPacks) | Measure-Object -Property TotalLicenses -Sum).Sum)
            $iss = [int]((@($Summary.InstalledRdsCalPacks) | Measure-Object -Property IssuedLicenses -Sum).Sum)
            $packInfo = " Kurulu RDS CAL havuzu: $iss/$tot kullanimda."
        }
        [void]$sb.Append("<div class='rec'><b>RDS CAL:</b> Tespit edilen mod = <b>$(ConvertTo-HtmlSafe $Summary.RdsMode)</b>. Tavsiye = <b>$(ConvertTo-HtmlSafe $Summary.Rds.RecommendedType)</b> &times; <b>$($Summary.Rds.RecommendedCount)</b>. <span class='muted'>$(ConvertTo-HtmlSafe $Summary.Rds.Note)</span>$packInfo</div>")
    } else {
        [void]$sb.Append("<div class='rec'><b>RDS CAL:</b> Sistemde RD Session Host rolu tespit edilmedi; RDS CAL gerekmiyor gorunuyor.</div>")
    }
    if (@($Summary.SqlPaid).Count -gt 0) {
        $sqlPackTotal = [int]((@($Summary.SqlPaid) | Measure-Object -Property PerCore_Packs2 -Sum).Sum)
        [void]$sb.Append("<div class='rec'><b>SQL Server:</b> $(@($Summary.SqlPaid).Count) ucretli instance. Per-Core modelinde toplam ~<b>$sqlPackTotal</b> adet 2-core paketi (min. soket basina 4 core). Standard instance'lar icin az/bilinen kullanici sayisinda Server+CAL modeli degerlendirilebilir.</div>")
    } else {
        [void]$sb.Append("<div class='rec'><b>SQL Server:</b> Ucretli SQL instance tespit edilmedi (yalnizca Express/Developer veya hic yok).</div>")
    }

    # --- Windows Server table ---
    [void]$sb.Append("<h2>Windows Server Envanteri &amp; Core Lisans</h2>")
    [void]$sb.Append("<table><tr><th>Sunucu</th><th>Isletim Sistemi</th><th>Tur</th><th>Soket</th><th>Fiziksel Core</th><th>Mantiksal</th><th>Lisanslanabilir Core</th><th>2-Core Paket</th><th>Edition</th></tr>")
    foreach ($r in $Summary.ServerLicenseRows) {
        $typePill = if ($r.IsVirtual) { "<span class='pill vm'>Sanal</span>" } else { "<span class='pill phys'>Fiziksel</span>" }
        $dcPill = if ($r.IsDC) { " <span class='pill dc'>DC</span>" } else { '' }
        [void]$sb.Append("<tr><td>$(ConvertTo-HtmlSafe $r.ComputerName)$dcPill</td><td>$(ConvertTo-HtmlSafe $r.OS)</td><td>$typePill</td><td>$($r.Sockets)</td><td>$($r.PhysicalCores)</td><td class='muted'>$($r.LogicalCores)</td><td><b>$($r.LicensableCores)</b></td><td>$($r.CorePacks_2Core)</td><td>$(ConvertTo-HtmlSafe $r.RecommendedEdition)</td></tr>")
    }
    [void]$sb.Append("</table>")
    [void]$sb.Append("<div class='sub' style='margin-top:8px'>Not: Sanal sunucularda gosterilen core, atanmis vCPU'dur. Microsoft kurallarinda Windows Server, VM'in uzerinde calistigi <b>fiziksel host'un tum core'lari</b> lisanslanarak (Datacenter = sinirsiz VM, Standard = lisans seti basina 2 VM) lisanslanir. Host&#8594;VM eslesmesi icin hipervizor envanteri (Get-VM / PowerCLI) gerekir.</div>")

    # --- SQL table ---
    if (@($Summary.SqlRows).Count -gt 0) {
        [void]$sb.Append("<h2>SQL Server Envanteri &amp; Lisans</h2>")
        [void]$sb.Append("<table><tr><th>Sunucu</th><th>Instance</th><th>Edition</th><th>Surum</th><th>Fiziksel Core</th><th>Per-Core (2-core paket)</th><th>Server+CAL?</th><th>Tavsiye Model</th></tr>")
        foreach ($r in $Summary.SqlRows) {
            $cls = switch ($r.EditionClass) { 'Enterprise' { 'ent' } 'Free' { 'free' } default { 'paid' } }
            $clsLbl = switch ($r.EditionClass) { 'Free' { 'Ucretsiz' } 'Enterprise' { 'Enterprise' } default { 'Ucretli' } }
            $spc = if ($r.ServerPlusCAL) { 'Evet (Standard)' } else { 'Hayir' }
            [void]$sb.Append("<tr><td>$(ConvertTo-HtmlSafe $r.ComputerName)</td><td>$(ConvertTo-HtmlSafe $r.Instance)</td><td>$(ConvertTo-HtmlSafe $r.Edition) <span class='pill $cls'>$clsLbl</span></td><td class='muted'>$(ConvertTo-HtmlSafe $r.Version)</td><td>$($r.PhysicalCores)</td><td><b>$($r.PerCore_Packs2)</b></td><td>$spc</td><td>$(ConvertTo-HtmlSafe $r.RecommendedModel)</td></tr>")
        }
        [void]$sb.Append("</table>")
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
        [void]$sb.Append("<div class='sub'>Bu sunuculara erisilemedigi icin lisans hesabina dahil edilmediler. WinRM (5985) / RPC (135) ve guvenlik duvari ayarlarini kontrol edin.</div>")
    }

    # --- Footer / sources ---
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

    $p1 = Join-Path $Directory "Servers_$Stamp.csv"
    $Summary.ServerLicenseRows | Select-Object ComputerName, OS, Type, Sockets, PhysicalCores, LogicalCores, LicensableCores, CorePacks_2Core, RecommendedEdition, IsDC |
        Export-Csv -Path $p1 -NoTypeInformation -Encoding UTF8
    $files += $p1

    if (@($Summary.SqlRows).Count -gt 0) {
        $p2 = Join-Path $Directory "SQL_$Stamp.csv"
        $Summary.SqlRows | Select-Object ComputerName, Instance, Edition, EditionClass, Version, PhysicalCores, PerCore_Cores, PerCore_Packs2, ServerPlusCAL, RecommendedModel |
            Export-Csv -Path $p2 -NoTypeInformation -Encoding UTF8
        $files += $p2
    }

    # Ozet CSV
    $p3 = Join-Path $Directory "Summary_$Stamp.csv"
    [PSCustomObject]@{
        ToplamSunucu          = $Summary.Inventory.Domain.ServerCount
        ErisilenSunucu        = $Summary.ReachableCount
        FizikselLisansCore_WS = $Summary.PhysicalCoreTotal
        FizikselCorePaket2    = $Summary.PhysicalPackTotal
        SanalCoreToplam       = $Summary.VirtualCoreTotal
        EtkinKullanici        = $Summary.Inventory.Domain.EnabledUserCount
        IsIstasyonu           = $Summary.Inventory.Domain.WorkstationCount
        CAL_Tavsiye           = "$($Summary.Cal.RecommendedType) x $($Summary.Cal.RecommendedCount)"
        RDS_Mod               = $Summary.RdsMode
        UcretliSQLInstance    = @($Summary.SqlPaid).Count
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

