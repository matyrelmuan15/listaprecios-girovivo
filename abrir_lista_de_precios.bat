@echo off
rem abrir_lista_de_precios.bat
rem ---------------------------
rem Doble clic para abrir la interfaz de escritorio del proyecto sin dejar
rem una ventana de consola abierta de fondo.
rem
rem Este .bat se puede usar directamente, o como destino de un acceso
rem directo en el Escritorio (ver crear_acceso_directo.ps1).

setlocal
cd /d "%~dp0"

rem Si en algun momento se suma un entorno virtual propio del proyecto
rem (.venv), se usa ese Python en vez del del sistema.
if exist "%~dp0.venv\Scripts\pythonw.exe" (
    start "" "%~dp0.venv\Scripts\pythonw.exe" "scripts\interfaz.py"
    goto :fin
)

where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "scripts\interfaz.py"
) else (
    start /min "" python "scripts\interfaz.py"
)

:fin
endlocal
