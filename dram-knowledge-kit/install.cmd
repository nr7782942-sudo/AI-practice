@echo off
setlocal DisableDelayedExpansion
if "%~1"=="" goto usage
if "%~2"=="" goto usage
if not "%~3"=="" goto usage
where py.exe >nul 2>nul
if errorlevel 1 goto check_python
py -3 --version >nul 2>nul
if not errorlevel 1 goto use_py
:check_python
where python.exe >nul 2>nul
if errorlevel 1 goto missing_python
python --version >nul 2>nul
if not errorlevel 1 goto use_python
:missing_python
echo ERROR: Approved Python 3.10 or newer is required. Nothing was installed.
exit /b 2
:use_py
py -3 "%~dp0install.py" --project "%~1" --vault "%~2"
exit /b %errorlevel%
:use_python
python "%~dp0install.py" --project "%~1" --vault "%~2"
exit /b %errorlevel%
:usage
echo Usage: install.cmd "C:\DRAM-Agent" "D:\Obsidian\DRAM"
echo Preview: py -3 "%~dp0install.py" --project "C:\DRAM-Agent" --vault "D:\Obsidian\DRAM" --dry-run
echo Nothing was installed.
exit /b 2
