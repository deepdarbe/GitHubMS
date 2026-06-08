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
