@echo off
title 🛠️ Build & Deploy analyse_pdf.exe

echo ==========================================
echo        🛠️  Outil de build : analyse_pdf
echo ==========================================

:: [1/4] Nettoyage
echo [1/4] Nettoyage des anciens fichiers...
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
echo    ✅ Fichiers nettoyés

:: [2/4] Compilation
echo [2/4] Compilation en cours...
pyinstaller --onefile --console ^
--add-data "dist_bundle_ressources;dist_bundle_ressources" ^
analyse_pdf.py

if errorlevel 1 (
    echo ❌ Erreur de compilation
    pause
    exit /b
)

echo    ✅ Compilation terminée !

:: [3/4] Copie automatique dans AnalysePDF
set TARGET=%USERPROFILE%\Documents\AnalysePDF
if not exist "%TARGET%" mkdir "%TARGET%"
copy /Y "dist\analyse_pdf.exe" "%TARGET%\analyse_pdf.exe"
echo    ✅ Copie terminée dans %TARGET%

:: [4/4] Mise à jour GitHub
echo [4/4] Push GitHub...
git add .
git commit -m "Build automatique : mise à jour de analyse_pdf.exe"
git push origin main
echo    ✅ Push terminé sur GitHub

echo.
echo ==========================================
echo 🚀 Build terminé avec succès !
echo L'exe est dispo dans %TARGET% et sur GitHub.
echo ==========================================
pause


