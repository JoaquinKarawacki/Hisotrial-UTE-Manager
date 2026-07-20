@echo off
cd /d "%~dp0"
setlocal

set "PYTHON_VERSION=3.12.0"
set "PYTHON_INSTALLER=%TEMP%\python_installer.exe"
set "PYTHON_EXE="

echo ================================================
echo Configuracion del entorno - Historial Facturas UTE
echo ================================================
echo.

echo [0/3] Verificando Python...
call :detect_python
if not defined PYTHON_EXE (
    echo Python no encontrado. Descargando instalador...
    curl -fL -o "%PYTHON_INSTALLER%" "https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-amd64.exe"
    if errorlevel 1 (
        echo ERROR: No se pudo descargar el instalador de Python.
        echo Verifica tu conexion a internet y vuelve a ejecutar setup.bat.
        pause
        exit /b 1
    )

    if not exist "%PYTHON_INSTALLER%" (
        echo ERROR: El instalador no se descargo correctamente.
        pause
        exit /b 1
    )

    echo Instalando Python...
    "%PYTHON_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1
    if errorlevel 1 (
        echo ERROR: La instalacion de Python fallo.
        del "%PYTHON_INSTALLER%" >nul 2>&1
        pause
        exit /b 1
    )
    del "%PYTHON_INSTALLER%" >nul 2>&1

    rem Refresca PATH en la sesion actual para detectar Python sin reiniciar.
    set "PATH=%LocalAppData%\Programs\Python\Python312\;%LocalAppData%\Programs\Python\Python312\Scripts\;%PATH%"

    call :detect_python
    if not defined PYTHON_EXE (
        echo ERROR: Python se instalo, pero no pudo detectarse en esta sesion.
        echo Cierra esta ventana y ejecuta setup.bat de nuevo.
        pause
        exit /b 1
    )
)
echo Python OK. Usando: %PYTHON_EXE%
echo.

echo [1/3] Creando entorno virtual...
call %PYTHON_EXE% -m venv .venv
if errorlevel 1 (
    echo ERROR: No se pudo crear el entorno virtual.
    echo Asegurate de tener Python instalado y en el PATH.
    pause
    exit /b 1
)

echo.
echo [2/3] Instalando dependencias...
.venv\Scripts\python -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Fallo la instalacion de dependencias.
    pause
    exit /b 1
)

echo.
echo [3/3] Descargando navegador Chromium para Playwright...
.venv\Scripts\python -m playwright install chromium
if errorlevel 1 (
    echo ERROR: Fallo la descarga de Chromium.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  Listo! Ahora podes iniciar la app con:
echo  Iniciar.bat
echo ================================================
pause
exit /b 0

:detect_python
set "PYTHON_EXE="
where py >nul 2>&1
if not errorlevel 1 (
    py -3 --version >nul 2>&1
    if not errorlevel 1 set "PYTHON_EXE=py -3"
)

if not defined PYTHON_EXE (
    where python >nul 2>&1
    if not errorlevel 1 (
        python --version >nul 2>&1
        if not errorlevel 1 set "PYTHON_EXE=python"
    )
)

if not defined PYTHON_EXE (
    if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
        set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python312\python.exe"
    )
)

exit /b 0
