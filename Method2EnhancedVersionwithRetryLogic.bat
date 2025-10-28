@echo off
setlocal enabledelayedexpansion

echo ========================================
echo   Office 2019 Activation Script
echo ========================================
echo.

:: Admin check
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] This script requires Administrator privileges.
    echo Please right-click and "Run as administrator"
    pause
    exit /b
)

:: Find Office directory
set OFFICE_DIR=
if exist "%ProgramFiles%\Microsoft Office\Office16\" (
    set OFFICE_DIR="%ProgramFiles%\Microsoft Office\Office16"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\" (
    set OFFICE_DIR="%ProgramFiles(x86)%\Microsoft Office\Office16"
) else (
    echo [ERROR] Office 2016/2019 directory not found.
    pause
    exit /b
)

echo Changing to: !OFFICE_DIR!
cd /d !OFFICE_DIR!

:: Install licenses
echo.
echo Step 1: Installing volume licenses...
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms 2^>nul') do (
    echo Installing license: %%x
    cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
)

:: Configure activation
echo.
echo Step 2: Configuring KMS settings...
cscript //nologo ospp.vbs /setprt:1688 >nul
cscript //nologo ospp.vbs /unpkey:6MWKP >nul 2>&1
cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul
cscript //nologo ospp.vbs /sethst:e8.us.to >nul

:: Activation with retry
echo.
echo Step 3: Attempting activation...
set max_retries=3
set retry_count=0

:retry_activation
set /a retry_count+=1
echo Attempt !retry_count! of !max_retries!...
cscript //nologo ospp.vbs /act

if !retry_count! lss !max_retries! (
    echo.
    choice /c yn /n /t 10 /d y /m "Activation may need retry. Try again (Y) or exit (N)?"
    if !errorlevel! equ 1 goto retry_activation
)

echo.
echo ========================================
echo   Activation process completed!
echo ========================================
echo.
echo If you see error 0xC004F074:
echo - Check your internet connection
echo - KMS server might be busy
echo - Try running this script again later
echo.
echo Check activation status with: cscript ospp.vbs /dstatus
echo.
pause