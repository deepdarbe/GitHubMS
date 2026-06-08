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
