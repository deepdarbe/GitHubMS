# HANDOFF — MS License Feasibility

Bu dosya oturumlar arası devamlılık (handoff) içindir. **Yalnızca araç durumunu**
özetler; müşteri/fiyat verisi içermez (öyle veriler `output/` altında üretilir ve
git'e dâhil edilmez).

## Durum
- **Branch:** `claude/determined-dijkstra-hBkQB`
- **PR:** deepdarbe/GitHubMS #3 (draft)
- **Olgunluk:** Çalışır ve test edilmiş. 4 modül + standalone **parse-temiz**;
  PSScriptAnalyzer (Warning/Error) proje ayarlarıyla **temiz**; demo uçtan uca
  doğrulandı. PowerShell **5.1 + 7** uyumlu (ternary/`System.Web` yok, saf ASCII).

## Araç ne yapıyor
Windows AD domain ortamını **ajansız** tarayıp Microsoft lisans ihtiyacını tahmin
eder ve **Lisans İhtiyaç Matrisi** + HTML rapor + CSV üretir:
- Aktif sunucu/OS, fiziksel core (8/soket, 16/sunucu min.), fiziksel↔sanal ayrımı
- **Host-bazlı Windows** (Datacenter vs Standard; `-PhysicalHostsFile`)
- Windows CAL (User/Device), RDS CAL (Per-User/Device)
- SQL Server (Per-Core vs Server+CAL; VM'de tüm vCPU/min 4; ücretsiz edition'lar elenir)

## Dosyalar
```
MSLicenseFeasibility/
├─ Run-Feasibility.cmd                 # click-to-run launcher (Gerçek/Liste/Demo)
├─ MSLicenseFeasibility.ps1            # ana orkestratör (çok dosyalı giriş)
├─ MSLicenseFeasibility-Standalone.ps1 # tek dosya (iex/wget; build-standalone ile ÜRETİLİR)
├─ build-standalone.ps1                # standalone üreticisi
├─ servers.txt / hosts.csv             # opsiyonel girdiler (örnek)
├─ lib/ {LicensingEngine, Collectors, ReportWriter}.ps1
├─ PSScriptAnalyzerSettings.psd1
└─ output/                             # üretilen raporlar (gitignore: html/csv/xlsx)
```

## Çalıştırma
```powershell
# Klasörden:
.\MSLicenseFeasibility.ps1 -Demo -OpenReport                  # önizleme
.\MSLicenseFeasibility.ps1 -PhysicalHostsFile .\hosts.csv -SqlUserCount 50 -RdsUserCount 30 -OpenReport

# Uzaktan tek dosya (repo public; SHA-pinli URL önerilir):
[Net.ServicePointManager]::SecurityProtocol='Tls12'
& ([scriptblock]::Create((irm 'https://raw.githubusercontent.com/deepdarbe/GitHubMS/claude/determined-dijkstra-hBkQB/MSLicenseFeasibility/MSLicenseFeasibility-Standalone.ps1'))) -Demo -OpenReport
```
> Kaynak (`lib/*.ps1` / ana script) değişince `build-standalone.ps1` ile standalone
> yeniden üretilmeli (elle düzenlenmez).

## Açık uçlar / olası sonraki adımlar
- Windows Server **CAL için `-UserCount` override** (RDS/SQL'de var, WS CAL AD sayısını kullanıyor)
- Matris HTML'inin doğrudan **`.xlsx` export**'u
- **Host→VM otomatik eşleme** (Hyper-V `Get-VM` / VMware PowerCLI) ile tam host-bazlı doğruluk
- Teklif tarafında **USD/TRY kuru + TL kolonu**, firma/teklif no başlığı, satır iskontosu

## Notlar
- Müşteriye özel envanter/teklif çıktıları (`output/*.xlsx|csv|html`) **gitignore**'da;
  repoya/PR'a girmez. Alış (maliyet) fiyatları hiçbir commit'te yer almaz.
- Lisans eşik değerleri Microsoft resmi rehberlerinden alınmıştır (bkz. README "Kaynaklar").
- Araç bir **tahmin/fizibilite** aracıdır; nihai lisanslama yetkili Microsoft iş ortağı ile teyit edilmelidir.
