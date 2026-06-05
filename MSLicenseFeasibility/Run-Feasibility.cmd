@echo off
REM ============================================================
REM  Microsoft Lisans Fizibilite Araci - Click-to-Run Launcher
REM  Bu dosyaya cift tiklayarak araci calistirabilirsiniz.
REM ============================================================
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
cd /d "%~dp0"
title Microsoft Lisans Fizibilite Araci

REM PowerShell 5.1 (powershell.exe) her Windows'ta hazir gelir.
REM PowerShell 7 (pwsh) varsa onu tercih etmek isterseniz asagidaki
REM PSEXE satirini "pwsh" yapabilirsiniz.
set "PSEXE=powershell.exe"

echo.
echo ============================================================
echo   Microsoft Lisans Fizibilite Araci
echo ============================================================
echo.
echo  Calistirma modunu secin:
echo.
echo    [1] Gercek tarama   - Active Directory'deki tum sunucular
echo    [2] Liste ile       - servers.txt dosyasindaki sunucular
echo    [3] Demo / Onizleme - ornek veri (domain gerektirmez)
echo.
set "MODE="
set /p "MODE=Seciminiz (1/2/3) [3]: "
if "%MODE%"=="" set "MODE=3"

echo.
if "%MODE%"=="1" (
    echo [i] Active Directory uzerinden gercek tarama baslatiliyor...
    "%PSEXE%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0MSLicenseFeasibility.ps1" -OpenReport
) else if "%MODE%"=="2" (
    echo [i] servers.txt listesi taraniyor...
    "%PSEXE%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0MSLicenseFeasibility.ps1" -ComputerListFile "%~dp0servers.txt" -OpenReport
) else (
    echo [i] Demo / onizleme modu calistiriliyor...
    "%PSEXE%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0MSLicenseFeasibility.ps1" -Demo -OpenReport
)

echo.
echo ------------------------------------------------------------
echo  Bitti. Cikti klasoru: %~dp0output
echo ------------------------------------------------------------
echo.
pause
endlocal
