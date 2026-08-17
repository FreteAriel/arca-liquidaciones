@echo off
title Subir proyecto a GitHub
cd /d "%~dp0"
if not exist ".git" goto firsttime
git add .
git commit -m "Actualizar proyecto liquidaciones"
git push
echo.
echo [OK] Cambios subidos a GitHub.
pause
exit /b 0
:firsttime
git init
git branch -m main
git add .
git commit -m "Primer commit - Liquidaciones IIBB/COM/IVA"
echo.
set /p REPO_URL="URL del repo GitHub (ej: https://github.com/FreteAriel/arca-liquidaciones.git): "
git remote add origin %REPO_URL%
git push -u origin main
echo.
echo [OK] Proyecto subido a GitHub.
pause
