@echo off
setlocal

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% == 0 (
    echo Python is already installed.
    exit /b 0
)

REM Download Python installer
echo Downloading Python installer...
bitsadmin /transfer "DownloadPython" https://www.python.org/ftp/python/3.10.5/python-3.10.5-amd64.exe %temp%\python-installer.exe

REM Run the installer
echo Installing Python...
%temp%\python-installer.exe /quiet InstallAllUsers=1 PrependPath=1

REM Clean up
del %temp%\python-installer.exe

REM Verify installation
python --version
if %errorlevel% == 0 (
    echo Python installed successfully.
) else (
    echo Python installation failed.
    exit /b 1
)

endlocal
