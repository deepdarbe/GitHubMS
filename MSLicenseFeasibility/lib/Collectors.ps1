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
