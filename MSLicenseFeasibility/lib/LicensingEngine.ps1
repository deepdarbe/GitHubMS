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
