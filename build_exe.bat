@echo off
REM ─── Dağıtım Mektubu → LFT Otomasyon — .exe Derleme Betiği ─────────────────
REM Çalıştırmadan önce: pip install -r requirements.txt
REM Çıktı: dist\DagitimMektubuLFT.exe

echo.
echo [1/3] Bağımlılıklar kuruluyor...
pip install -r requirements.txt --quiet

echo [2/3] .exe derleniyor...
pyinstaller ^
  --onefile ^
  --windowed ^
  --name "DagitimMektubuLFT" ^
  --add-data "core;core" ^
  --add-data "utils;utils" ^
  --icon NONE ^
  app.py

echo [3/3] Temizlik...
rmdir /s /q build
del /q *.spec

echo.
echo ─────────────────────────────────────────────
echo  Hazir! dist\DagitimMektubuLFT.exe
echo  Bu dosyayi personele dagitabilirsiniz.
echo  Python kurulumu gerekmez.
echo ─────────────────────────────────────────────
pause
