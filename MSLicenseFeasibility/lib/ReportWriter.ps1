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
