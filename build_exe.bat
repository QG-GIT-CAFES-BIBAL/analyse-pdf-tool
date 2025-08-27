@echo off
echo === Rebuild analyse_pdf.exe ===
pyinstaller --onefile --console ^
--add-data "dist_bundle_ressources;dist_bundle_ressources" ^
analyse_pdf.py

echo.
echo ✅ Compilation terminée !
echo Le nouvel exe est dispo dans "dist\analyse_pdf.exe"
pause
