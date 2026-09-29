@echo off
cd /d "%~dp0"
echo.
echo ============================================
echo  Compilando DistribuidorHotelGUI.exe ...
echo ============================================
echo.

pyinstaller --noconfirm --clean --onefile --windowed ^
  --name "DistribuidorHotelGUI" ^
  --hidden-import demo ^
  --hidden-import customtkinter ^
  interfaz.py

echo.
echo ============================================
echo  Listo. El exe queda en dist\DistribuidorHotelGUI.exe
echo ============================================
pause