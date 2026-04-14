@echo off
:: ============================================================
::  BUILD — LicitaMonitor Soldesp
::  Ejecutar en el equipo Windows donde se compila.
::  Requiere Python 3.10+ y Google Chrome instalado.
:: ============================================================

setlocal EnableDelayedExpansion
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo.
echo  =====================================================
echo   LICITAMONITOR — Compilacion del ejecutable
echo  =====================================================
echo.

:: ── 1. Verificar Python ────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no encontrado. Instala Python 3.10+ desde python.org
    pause & exit /b 1
)
echo [OK] Python: & python --version

:: ── 2. Instalar dependencias ──────────────────────────────
echo.
echo [INFO] Instalando dependencias...
python -m pip install --upgrade pip --quiet
python -m pip install --upgrade ^
    pyinstaller ^
    selenium ^
    openpyxl ^
    webdriver-manager ^
    requests ^
    certifi ^
    urllib3 ^
    Pillow ^
    --quiet

if errorlevel 1 (
    echo [ERROR] Fallo al instalar dependencias.
    pause & exit /b 1
)
echo [OK] Dependencias listas.

:: ── 3. Limpiar builds anteriores ──────────────────────────
echo.
echo [INFO] Limpiando builds anteriores...
for %%D in (build dist __pycache__ LicitaMonitor.build LicitaMonitor.dist) do (
    if exist "%%D" rmdir /s /q "%%D"
)
echo [OK] Limpieza completada.

:: ── 4. Compilar ───────────────────────────────────────────
echo.
echo [INFO] Compilando LicitaMonitor.exe ...
echo        (puede tardar 3-6 minutos)
echo.
pyinstaller LicitaMonitor.spec --noconfirm --clean
if errorlevel 1 (
    echo.
    echo [ERROR] La compilacion fallo. Revisa los mensajes anteriores.
    pause & exit /b 1
)

:: ── 5. Verificar salida ───────────────────────────────────
if not exist "LicitaMonitor.exe" (
    echo [WARN] LicitaMonitor.exe no esta en la carpeta raiz.
    echo        Buscando en dist\...
    if exist "dist\LicitaMonitor.exe" (
        copy /y "dist\LicitaMonitor.exe" "LicitaMonitor.exe" >nul
        echo [OK] Copiado desde dist\
    ) else (
        echo [ERROR] No se encontro LicitaMonitor.exe
        pause & exit /b 1
    )
)

:: ── 6. Crear carpeta de distribucion ──────────────────────
echo.
echo [INFO] Creando carpeta de distribucion...
set "DIST_DIR=%SCRIPT_DIR%Distribucion_LicitaMonitor"
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
mkdir "%DIST_DIR%"

copy /y "LicitaMonitor.exe"  "%DIST_DIR%\LicitaMonitor.exe"  >nul
copy /y "config.json"        "%DIST_DIR%\config.json"        >nul
if exist "logo.png"   copy /y "logo.png"   "%DIST_DIR%\logo.png"   >nul
if exist "icon.ico"   copy /y "icon.ico"   "%DIST_DIR%\icon.ico"   >nul
if exist "INSTALACION_EQUIPO_NUEVO.txt" (
    copy /y "INSTALACION_EQUIPO_NUEVO.txt" "%DIST_DIR%\INSTALACION.txt" >nul
)

echo.
echo  =====================================================
echo   COMPILACION EXITOSA
echo  =====================================================
echo.
echo  Ejecutable : %DIST_DIR%\LicitaMonitor.exe
echo  Config     : %DIST_DIR%\config.json
echo.
echo  Copia la carpeta "Distribucion_LicitaMonitor" al equipo destino.
echo  Edita config.json si las credenciales necesitan cambios.
echo.
pause
