@echo off
chcp 65001 >nul
title Subir proyecto a GitHub

echo ============================================================
echo   SUBIR PROYECTO A GITHUB — Liquidaciones ARCA/ARBA
echo ============================================================
echo.

cd /d "%~dp0"

:: Verificar si git ya fue inicializado
if exist ".git" (
    echo [OK] Repositorio git ya existe, actualizando...
    git add .
    git commit -m "Actualizar proyecto liquidaciones"
    git push
    echo.
    echo [OK] Cambios subidos a GitHub.
    pause
    exit /b 0
)

:: Primera vez: inicializar git y conectar con GitHub
echo [PASO 1] Inicializando repositorio git...
git init
git branch -m main

echo.
echo [PASO 2] Agregando todos los archivos...
git add .
git commit -m "Primer commit — Libros IVA + Liquidaciones IIBB/COM/IVA"

echo.
echo ============================================================
echo   AHORA NECESITAS CREAR EL REPO EN GITHUB:
echo.
echo   1. Abri https://github.com/new
echo   2. Nombre del repo: arca-liquidaciones
echo   3. Dejalo en PRIVADO (Private)
echo   4. NO marques "Add README"
echo   5. Hace clic en "Create repository"
echo   6. Copiá el comando que aparece:
echo      git remote add origin https://github.com/FreteAriel/arca-liquidaciones.git
echo      (te lo va a mostrar GitHub en pantalla)
echo ============================================================
echo.
set /p REPO_URL="Pega aqui la URL del repo (ej: https://github.com/FreteAriel/arca-liquidaciones.git): "

git remote add origin %REPO_URL%
git push -u origin main

echo.
echo ============================================================
echo   [OK] PROYECTO SUBIDO A GITHUB
echo   Ahora podes ir a Railway y conectar ese repo.
echo ============================================================
pause
