@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Ek satis paneli olusturuluyor...
python build_ek_satis.py ek_satis.xlsx
if errorlevel 1 (
  echo.
  echo HATA olustu, GitHub a yukleme yapilmadi.
  pause
  exit /b 1
)
echo GitHub a yukleniyor...
git add index.html
git commit -m "Ek satis guncelleme %date% %time%"
git push
echo.
echo Bitti. Birkac dakika icinde yayinda olur.
pause
