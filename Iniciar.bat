@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
	echo ================================================
	echo ERROR: No se encontro el entorno virtual ^(.venv^)
	echo Ejecuta setup.bat para instalar dependencias.
	echo ================================================
	pause
	exit /b 1
)

".venv\Scripts\python.exe" src\historial_manager.py
set EXIT_CODE=%ERRORLEVEL%

if not "%EXIT_CODE%"=="0" (
	echo.
	echo ================================================
	echo La aplicacion termino con error (codigo %EXIT_CODE%).
	echo Revisa el mensaje de error mostrado arriba.
	echo ================================================
	pause
)
