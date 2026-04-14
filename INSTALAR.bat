@echo off
:: ============================================================
::  LICITAMONITOR — Instalador para equipo nuevo
::  Soldesp · Portal Minero
::  Doble clic para instalar. Requiere internet.
:: ============================================================

setlocal EnableDelayedExpansion
title LicitaMonitor — Instalador Soldesp

echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║        LICITAMONITOR — INSTALADOR SOLDESP           ║
echo  ║             Portal Minero · v3.2                    ║
echo  ╚══════════════════════════════════════════════════════╝
echo.

:: Carpeta donde está este .bat (la del pendrive o donde se copió)
set "SRC=%~dp0"
cd /d "%SRC%"

:: Carpeta destino en el equipo
set "DEST=C:\LicitaMonitor"

:: ── 1. Verificar Google Chrome ─────────────────────────────────────────
echo [1/6] Verificando Google Chrome...
set "CHROME_OK=0"
for %%P in (
    "%ProgramFiles%\Google\Chrome\Application\chrome.exe"
    "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
    "%LocalAppData%\Google\Chrome\Application\chrome.exe"
) do (
    if exist %%P set "CHROME_OK=1"
)

if "%CHROME_OK%"=="0" (
    echo.
    echo  ┌─────────────────────────────────────────────────────┐
    echo  │  ATENCION: Google Chrome no está instalado.         │
    echo  │                                                     │
    echo  │  LicitaMonitor necesita Chrome para funcionar.      │
    echo  │  Instálalo desde: https://www.google.com/chrome/    │
    echo  │  y vuelve a ejecutar este instalador.               │
    echo  └─────────────────────────────────────────────────────┘
    echo.
    pause
    exit /b 1
)
echo  [OK] Google Chrome encontrado.

:: ── 2. Verificar / instalar Python ─────────────────────────────────────
echo.
echo [2/6] Verificando Python 3.10+...

set "PYTHON_OK=0"
for %%C in (python python3 py) do (
    %%C --version >nul 2>&1
    if not errorlevel 1 (
        for /f "tokens=2" %%V in ('%%C --version 2^>^&1') do (
            for /f "tokens=1,2 delims=." %%A in ("%%V") do (
                if %%A GEQ 3 if %%B GEQ 10 (
                    set "PYTHON_CMD=%%C"
                    set "PYTHON_OK=1"
                )
            )
        )
    )
)

if "%PYTHON_OK%"=="0" (
    echo  Python 3.10+ no encontrado. Descargando instalador...
    echo  (requiere conexion a internet — puede tardar unos minutos)
    echo.

    :: Descargar Python 3.12 silencioso
    set "PY_INSTALLER=%TEMP%\python_installer.exe"
    powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe' -OutFile '%PY_INSTALLER%'" 2>nul

    if not exist "%PY_INSTALLER%" (
        echo  [ERROR] No se pudo descargar Python. Verifica tu conexion a internet.
        echo  Descargalo manualmente desde https://www.python.org/ (Python 3.12)
        echo  y vuelve a ejecutar este instalador.
        pause & exit /b 1
    )

    echo  Instalando Python 3.12 (instalacion silenciosa)...
    "%PY_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_tcltk=1
    del /f "%PY_INSTALLER%" >nul 2>&1

    :: Refrescar PATH
    call refreshenv >nul 2>&1
    set "PYTHON_CMD=python"
    echo  [OK] Python instalado. Es posible que necesites cerrar y reabrir si falla.
) else (
    echo  [OK] Python encontrado: %PYTHON_CMD%
)

:: ── 3. Copiar archivos al equipo ────────────────────────────────────────
echo.
echo [3/6] Copiando archivos a %DEST%...

if not exist "%DEST%" mkdir "%DEST%"
if not exist "%DEST%\logs" mkdir "%DEST%\logs"
if not exist "%DEST%\Excel" mkdir "%DEST%\Excel"

copy /y "%SRC%LicitaMonitor.py"     "%DEST%\LicitaMonitor.py"     >nul
copy /y "%SRC%LicitaMonitor_GUI.py" "%DEST%\LicitaMonitor_GUI.py" >nul
if exist "%SRC%logo.png"   copy /y "%SRC%logo.png"   "%DEST%\logo.png"   >nul
if exist "%SRC%icon.ico"   copy /y "%SRC%icon.ico"   "%DEST%\icon.ico"   >nul

:: config.json: copiar SOLO si no existe ya en destino (no sobreescribir credenciales)
if not exist "%DEST%\config.json" (
    copy /y "%SRC%config.json" "%DEST%\config.json" >nul
    echo  [OK] config.json copiado. Editar credenciales si es necesario.
) else (
    echo  [OK] config.json ya existe — no sobreescrito (credenciales conservadas).
)

echo  [OK] Archivos copiados.

:: ── 4. Instalar dependencias Python ────────────────────────────────────
echo.
echo [4/6] Instalando dependencias Python...
echo       (selenium, openpyxl, webdriver-manager, Pillow...)

%PYTHON_CMD% -m pip install --upgrade pip --quiet --no-warn-script-location
%PYTHON_CMD% -m pip install --upgrade ^
    selenium ^
    openpyxl ^
    webdriver-manager ^
    requests ^
    certifi ^
    Pillow ^
    --quiet --no-warn-script-location

if errorlevel 1 (
    echo  [ERROR] Fallo al instalar dependencias. Verifica tu conexion a internet.
    pause & exit /b 1
)
echo  [OK] Dependencias instaladas.

:: ── 5. Crear lanzador .bat en destino ──────────────────────────────────
echo.
echo [5/6] Creando lanzador...

set "LAUNCHER=%DEST%\Abrir_LicitaMonitor.bat"
(
echo @echo off
echo cd /d "C:\LicitaMonitor"
echo start "" pythonw LicitaMonitor_GUI.py
) > "%LAUNCHER%"

:: Crear acceso directo en el Escritorio del usuario actual
set "SHORTCUT=%USERPROFILE%\Desktop\LicitaMonitor.lnk"
powershell -Command "$s=(New-Object -COM WScript.Shell).CreateShortcut('%SHORTCUT%'); $s.TargetPath='%LAUNCHER%'; $s.WorkingDirectory='C:\LicitaMonitor'; $s.Description='LicitaMonitor Soldesp'; $s.Save()" >nul 2>&1

if exist "%SHORTCUT%" (
    echo  [OK] Acceso directo creado en el Escritorio.
) else (
    echo  [OK] Lanzador creado en %LAUNCHER%
    echo      (acceso directo en escritorio no disponible — usa el .bat directamente)
)

:: ── 6. Verificar instalacion ────────────────────────────────────────────
echo.
echo [6/6] Verificando instalacion...

%PYTHON_CMD% -c "import selenium, openpyxl, webdriver_manager; print('OK')" >nul 2>&1
if errorlevel 1 (
    echo  [WARN] Alguna dependencia no se importo correctamente.
    echo         Intenta ejecutar de nuevo o contacta soporte.
) else (
    echo  [OK] Todo listo.
)

:: ── Resumen final ────────────────────────────────────────────────────────
echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║              INSTALACION COMPLETADA                 ║
echo  ╠══════════════════════════════════════════════════════╣
echo  ║  Carpeta : C:\LicitaMonitor\                        ║
echo  ║  Lanzar  : Doble clic en "LicitaMonitor"            ║
echo  ║            en el Escritorio                         ║
echo  ║                                                     ║
echo  ║  IMPORTANTE: edita C:\LicitaMonitor\config.json     ║
echo  ║  si necesitas cambiar usuario o contraseña.         ║
echo  ╚══════════════════════════════════════════════════════╝
echo.

set /p "_=  Presiona Enter para abrir LicitaMonitor ahora... "
start "" pythonw "%DEST%\LicitaMonitor_GUI.py"
exit /b 0
