# Microsoft Lisans Fizibilite Aracı (MS License Feasibility)

Windows **Active Directory domain** ortamında, Microsoft lisans ihtiyacını
netleştirmek için kullanılan **ajansız (agentless), click-to-run** bir
PowerShell aracı. Domaindeki sunucuları ve istemcileri tarayarak Windows
Server (core), Windows CAL, RDS CAL ve SQL Server lisans ihtiyacını **tahmin
eder** ve bir HTML yönetici raporu + CSV çıktıları üretir.

> ⚠️ **Uyarı:** Bu bir **fizibilite / tahmin** aracıdır. Lisanslama kuralları
> Microsoft'un resmi rehberlerine dayanır; ancak nihai lisans ihtiyacı
> Software Assurance, sözleşme tipi, sanallaştırma hakları ve host→VM
> eşleşmesi gibi etkenlere göre değişir. **Satın alma öncesi yetkili bir
> Microsoft lisans uzmanıyla teyit edin.**

---

## Ne topluyor?

| Alan | Toplanan veri | Yöntem |
|------|---------------|--------|
| **Sunucular** | Aktif sunucu sayısı, işletim sistemi (caption/sürüm/build) | AD `Get-ADComputer` + `Win32_OperatingSystem` |
| **Core** | Soket sayısı, **fiziksel core** (lisanslanabilir core matematiği), mantıksal işlemci | `Win32_Processor` (hyper-threading sayılmaz) |
| **Fiziksel/Sanal** | Fiziksel mi VM mi (Datacenter vs Standard kararı için) | `Win32_ComputerSystem` üretici/model imzası |
| **SQL Server** | Instance adları, **edition** (Express/Developer = ücretsiz, Standard/Enterprise = ücretli), sürüm | Uzak registry + SQL WMI provider |
| **RDS** | RD Session Host rolü, lisans modu (**Per-User / Per-Device**), kurulu RDS CAL paketleri | `Win32_TerminalServiceSetting`, GPO/RCM registry, `Win32_TSLicenseKeyPack` |
| **CAL** | Etkin kullanıcı sayısı + iş istasyonu/PC sayısı (User vs Device CAL) | AD LDAP bit filtresi |

---

## Gereksinimler

- **Windows** (domaine üye bir yönetici makinesi veya bir Domain Controller)
- **Windows PowerShell 5.1** (her Windows'ta hazır gelir) veya **PowerShell 7+**
- **RSAT ActiveDirectory modülü** (`Get-ADComputer`, `Get-ADUser` için)
  - Sunucuda: `Install-WindowsFeature RSAT-AD-PowerShell`
  - Windows 10/11'de: "RSAT: Active Directory" isteğe bağlı özelliği
  - _AD modülü yoksa_ `-ComputerListFile` ile manuel sunucu listesi kullanılabilir.
- Uzak sunuculara erişim için **yönetici yetkisi** ve şu portlardan en az biri:
  - **WinRM/WS-MAN (TCP 5985)** — tercih edilen
  - **RPC/DCOM (TCP 135 + dinamik)** — otomatik fallback
- Hedef sunucularda **RemoteRegistry** servisi (SQL/RDS registry okuması için; çoğu ortamda açıktır)

---

## Çalıştırma (Click-to-Run)

En kolay yol: **`Run-Feasibility.cmd`** dosyasına çift tıklayın. Bir menü açılır:

```
   [1] Gercek tarama   - Active Directory'deki tum sunucular
   [2] Liste ile       - servers.txt dosyasindaki sunucular
   [3] Demo / Onizleme - ornek veri (domain gerektirmez)
```

İlk kez denemek için **[3] Demo**'yu seçin — domain gerektirmez, örnek bir
rapor üretip açar.

### PowerShell ile doğrudan

```powershell
# Tüm domaini tara, rapor üret ve aç
.\MSLicenseFeasibility.ps1 -OpenReport

# Belirli bir OU ile sınırla
.\MSLicenseFeasibility.ps1 -SearchBase "OU=Servers,DC=contoso,DC=local"

# Manuel sunucu listesi + alternatif kimlik
.\MSLicenseFeasibility.ps1 -ComputerListFile .\servers.txt -Credential (Get-Credential)

# Domain olmadan örnek rapor (önizleme)
.\MSLicenseFeasibility.ps1 -Demo -OpenReport
```

### Parametreler

| Parametre | Açıklama | Varsayılan |
|-----------|----------|------------|
| `-Demo` | Sentetik örnek veriyle rapor üretir (domain gerektirmez) | — |
| `-ComputerListFile <yol>` | AD keşfi yerine dosyadaki sunucuları tarar | — |
| `-SearchBase <DN>` | AD sorgularını bir OU ile sınırlar | tüm domain |
| `-StaleDays <n>` | Bu günden eski (LastLogon) nesneler "pasif" sayılır | 90 |
| `-Credential <pscred>` | Uzak bağlantılar için alternatif kimlik | mevcut kullanıcı |
| `-OutputDir <yol>` | Rapor klasörü | `.\output` |
| `-PerHostTimeoutSec <n>` | Sunucu başına CIM zaman aşımı (sn) | 20 |
| `-NoHtml` | Yalnızca CSV üret | — |
| `-OpenReport` | Rapor oluşunca tarayıcıda aç | — |

---

## Tek dosya ile uzaktan çalıştırma (iex / iwr / wget)

Çok dosyalı sürüm `lib\*.ps1` dosyalarını dot-source ettiği için
`iex (irm ...)` ile **doğrudan çalışmaz** (`$PSScriptRoot` bellekte boştur).
Bunun için tüm modüller tek dosyada toplanmış
**`MSLicenseFeasibility-Standalone.ps1`** üretilir (`build-standalone.ps1` ile).

> ⚠️ **Güvenlik:** İnternetten indirip belleğe alarak betik çalıştırmak (iex),
> o betiğin tüm kodunu çalıştırır. Yalnızca **güvendiğiniz** kaynaktan ve
> tercihen **commit SHA'sına sabitlenmiş** URL'den çalıştırın. Aşağıdaki
> örnekler reponun **public** olduğunu varsayar; repo private ise en alttaki
> nota bakın.

PowerShell 5.1+ (domaine üye, yönetici makinesi):

```powershell
# Eski Windows / PowerShell 5.1'de GitHub'a bağlanmak için önce TLS 1.2'yi açın:
[Net.ServicePointManager]::SecurityProtocol = 'Tls12'

# Parametreli (ÖNERİLEN) — belleğe indirip çalıştırır:
& ([scriptblock]::Create((irm 'RAW_URL'))) -Demo -OpenReport     # önizleme
& ([scriptblock]::Create((irm 'RAW_URL'))) -OpenReport           # gerçek tarama

# iex ile (parametre geçilmez; varsayılan = gerçek AD taraması):
iex (irm 'RAW_URL')

# wget/iwr ile indir, sonra çalıştır:
iwr 'RAW_URL' -OutFile "$env:TEMP\MSLF.ps1"
powershell -ExecutionPolicy Bypass -File "$env:TEMP\MSLF.ps1" -OpenReport
```

**`RAW_URL`** (bu geliştirme dalı):
```
https://raw.githubusercontent.com/deepdarbe/GitHubMS/claude/determined-dijkstra-hBkQB/MSLicenseFeasibility/MSLicenseFeasibility-Standalone.ps1
```
`master`'a merge sonrası:
```
https://raw.githubusercontent.com/deepdarbe/GitHubMS/master/MSLicenseFeasibility/MSLicenseFeasibility-Standalone.ps1
```

- Bellekten (iex) çalıştırıldığında çıktılar **bulunduğunuz dizindeki**
  `output\` klasörüne yazılır.
- **Private repo** ise ham URL bir PAT (token) ister:
  ```powershell
  $h = @{ Authorization = 'token <PAT>' }
  & ([scriptblock]::Create((irm -Headers $h 'RAW_URL'))) -Demo
  ```
  Alternatif: dosyayı tarayıcıdan indirin ya da çok dosyalı klasörü kopyalayıp
  `Run-Feasibility.cmd` kullanın.

### Standalone'u yeniden üretme
Kaynak (`lib\*.ps1` veya ana script) değiştiğinde:
```powershell
.\build-standalone.ps1
```
`MSLicenseFeasibility-Standalone.ps1` **otomatik üretilir — elle düzenlemeyin.**

---

## Çıktılar

`output\` klasörüne zaman damgalı olarak yazılır:

- **`LisansFizibilite_<tarih>.html`** — Yönetici özeti, kart göstergeleri ve
  detay tablolarıyla tek dosyalık HTML rapor.
- **`Servers_<tarih>.csv`** — Sunucu bazında OS / soket / core / edition.
- **`SQL_<tarih>.csv`** — SQL instance / edition / lisans modeli.
- **`Summary_<tarih>.csv`** — Tek satırlık genel özet (satın alma tablosu için).

---

## Kodlanmış lisans kuralları (Microsoft resmi)

| Kural | Değer |
|-------|-------|
| Windows Server — soket başına min. core | **8** |
| Windows Server — sunucu başına min. core | **16** |
| Windows Server — core paket boyutu | **2-core** |
| Windows Server Standard — lisans seti başına VM | **2** (daha fazlası için "stacking") |
| Windows Server Datacenter — VM | **Sınırsız** (yoğun sanallaştırmada önerilir) |
| Windows/RDS CAL sürümü | CAL sürümü ≥ sunucu sürümü |
| RDS CAL | Windows CAL'a **ek**; Per-User / Per-Device |
| SQL Server — soket başına min. core | **4** |
| SQL Server — core paket boyutu | **2-core** |
| SQL Server+CAL modeli | Yalnızca Standard (yeni anlaşmalarda) |

**Kararlar nasıl veriliyor?**
- **Lisanslanabilir core** = `max(16, Σ soket × max(8, soketteki core))` — ham
  core değil, minimumlar uygulanmış değer. SQL için minimum soket başına 4.
- **Standard vs Datacenter** = sunucu fiziksel + sanallaştırma host'u ise VM
  yoğunluğuna göre (varsayılan eşik **14 VM**); OS caption'da "Datacenter"
  geçiyorsa doğrudan Datacenter.
- **User vs Device CAL** = kullanıcı ≤ cihaz ise User CAL, değilse Device CAL.
- **SQL Per-Core vs Server+CAL** = Enterprise → Per-Core; Standard + bilinen/az
  kullanıcı → Server+CAL değerlendirilir.

### Kaynaklar
- [Core-based licensing models — Microsoft Licensing Guidance](https://www.microsoft.com/licensing/guidance/Core-based-licensing-models)
- [Windows Server 2025 / 2022 Licensing Guides](https://www.microsoft.com/en-us/licensing/product-licensing/windows-server)
- [Client Access License (CAL)](https://www.microsoft.com/en-us/licensing/product-licensing/client-access-license)
- [License Remote Desktop Services with CALs — learn.microsoft.com](https://learn.microsoft.com/windows-server/remote/remote-desktop-services/rds-client-access-license)
- [SQL Server Licensing Guidance](https://www.microsoft.com/licensing/guidance/SQL)

---

## Mimari

```
MSLicenseFeasibility/
├── Run-Feasibility.cmd                 # Click-to-run launcher (menü)
├── MSLicenseFeasibility.ps1            # Ana orkestratör (çok dosyalı giriş)
├── MSLicenseFeasibility-Standalone.ps1 # Tek dosya (iex/wget için; ÜRETİLİR)
├── build-standalone.ps1                # Standalone üreticisi
├── servers.txt                         # (opsiyonel) manuel sunucu listesi
├── lib/
│   ├── LicensingEngine.ps1             # Lisans hesaplama motoru (kurallar)
│   ├── Collectors.ps1                  # AD + CIM/WMI veri toplama (agentless)
│   └── ReportWriter.ps1                # HTML + CSV rapor üretimi
├── output/                             # Üretilen raporlar (gitignore)
└── PSScriptAnalyzerSettings.psd1
```

- **Ajansız:** Hedef sunuculara kurulum gerektirmez; CIM (WS-MAN) üzerinden,
  başarısız olursa DCOM'a düşerek sorgular.
- **Güvenli:** Yalnızca **okuma** yapar (hiçbir `Set-*` / değişiklik yok). Her
  uzak çağrı erişilebilirlik testinden geçer ve hata durumunda sunucu ayrı
  "erişilemeyen" listesinde raporlanır.
- **Uyumlu:** Windows PowerShell 5.1 ve PowerShell 7 ile çalışır (ternary /
  `System.Web` bağımlılığı yoktur).

---

## Bilinen kısıtlamalar

- **Host→VM eşleşmesi:** Ajansız in-guest sorgu, bir VM'in çalıştığı fiziksel
  host'un core sayısını göremez. Windows Server, VM'in üzerinde çalıştığı
  **fiziksel host** lisanslanarak karşılanır. Rapor sanal sunucuların atanmış
  vCPU'sunu ayrı gösterir; tam host bazlı hesap için hipervizör envanteri
  (`Get-VM` / VMware PowerCLI) ile zenginleştirilmelidir.
- **Software Assurance:** Per-VM lisanslama ve SQL sınırsız sanallaştırma gibi
  SA gerektiren senaryolar varsayılan olarak hesaba katılmaz (SA'sız / perpetual
  varsayılır).
- **CAL sayıları**, etkin AD kullanıcı/iş istasyonu sayılarından türetilir;
  servis hesapları veya harici/anonim kullanıcılar için ince ayar gerekebilir.
- Büyük ortamlarda tarama sıralı yapılır; çok sayıda sunucuda süre uzayabilir.

---

## Lisans / sorumluluk reddi

Bu araç envanter ve **tahmini** lisans hesabı sağlar. Üretilen rakamlar
bağlayıcı değildir ve resmi lisans danışmanlığı yerine geçmez.
