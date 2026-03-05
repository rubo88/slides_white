@echo off
cd /d "%~dp0"

:: Get current Python version
for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYVER=%%v

:: Check if lib was built for this Python version
if not exist "%~dp0lib\.installed" goto :install
set /p INSTALLED_VER=<"%~dp0lib\.installed"
if not "%PYVER%"=="%INSTALLED_VER%" goto :install
goto :run

:install
echo Installing required packages for %PYVER%...
if exist "%~dp0lib" rd /s /q "%~dp0lib"
mkdir "%~dp0lib"
python -m pip install python-pptx lxml --target "%~dp0lib" --quiet
echo %PYVER%>"%~dp0lib\.installed"
echo Done.

:run
set PYTHONPATH=%~dp0lib;%PYTHONPATH%
python fix_chart_colors.py
pause
