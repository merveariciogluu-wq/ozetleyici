@echo off
chcp 65001 >nul
title Logo Updater - Kurulum

echo ============================================
echo   Logo Updater - Bagimlilik Kurulumu
echo ============================================
echo.

:: Python var mı kontrol et
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo HATA: Python bulunamadi!
    echo Lutfen https://www.python.org adresinden Python 3.8+ indirip kurun.
    echo Kurulum sirasinda "Add Python to PATH" secenegini isaretleyin.
    echo.
    pause
    exit /b 1
)

echo Python bulundu:
python --version
echo.

:: requirements.txt bu klasorde mi?
if not exist "%~dp0requirements.txt" (
    echo HATA: requirements.txt dosyasi bulunamadi.
    echo Bu script logo_updater klasorunde calistirilmalidir.
    echo.
    pause
    exit /b 1
)

echo openpyxl kuruluyor...
echo.
python -m pip install -r "%~dp0requirements.txt"

if %errorlevel% equ 0 (
    echo.
    echo ============================================
    echo   Kurulum basariyla tamamlandi!
    echo   Artik main.py'yi calistirabilirsiniz:
    echo     python main.py DOSYA.xlsx
    echo ============================================
) else (
    echo.
    echo HATA: Kurulum basarisiz oldu.
    echo pip guncel degil olabilir. Deneyiniz:
    echo   python -m pip install --upgrade pip
    echo   python -m pip install -r requirements.txt
)

echo.
pause
