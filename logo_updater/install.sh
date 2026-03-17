#!/bin/bash
# Logo Updater - Bağımlılık Kurulum Scripti

echo "============================================"
echo "  Logo Updater - Bağımlılık Kurulumu"
echo "============================================"
echo

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Python komutu belirle (python3 veya python)
if command -v python3 &>/dev/null; then
    PYTHON=python3
    PIP="python3 -m pip"
elif command -v python &>/dev/null; then
    PYTHON=python
    PIP="python -m pip"
else
    echo "HATA: Python bulunamadı!"
    echo "Lütfen Python 3.8+ kurun: https://www.python.org"
    exit 1
fi

echo "Python bulundu: $($PYTHON --version)"
echo

# requirements.txt kontrolü
if [ ! -f "$SCRIPT_DIR/requirements.txt" ]; then
    echo "HATA: requirements.txt bulunamadı."
    exit 1
fi

echo "openpyxl kuruluyor..."
echo
$PIP install -r "$SCRIPT_DIR/requirements.txt"

if [ $? -eq 0 ]; then
    echo
    echo "============================================"
    echo "  Kurulum başarıyla tamamlandı!"
    echo "  Kullanım: python main.py DOSYA.xlsx"
    echo "============================================"
else
    echo
    echo "HATA: Kurulum başarısız oldu."
    echo "Manuel kurulum için: $PIP install openpyxl"
    exit 1
fi
