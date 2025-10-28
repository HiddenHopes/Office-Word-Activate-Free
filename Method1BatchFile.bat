@echo off
echo Running Office 2019 Activation Script...
echo.

:: Check if running as administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo This script requires Administrator privileges.
    echo Please right-click and "Run as administrator"
    pause
    exit /b
)

:: Navigate to Office directory
echo Changing to Office directory...
cd /d "%ProgramFiles%\Microsoft Office\Office16"
if %errorLevel% neq 0 (
    echo Office directory not found. Trying alternative path...
    cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
    if %errorLevel% neq 0 (
        echo Office 2016/2019 not found in standard locations.
        pause
        exit /b
    )
)

:: Install volume license
echo Installing volume license...
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"

:: Configure KMS and activate
echo Configuring KMS server...
cscript ospp.vbs /setprt:1688 
cscript ospp.vbs /unpkey:6MWKP >nul 
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP 
cscript ospp.vbs /sethst:e8.us.to 

echo Attempting activation...
cscript ospp.vbs /act

echo.
echo If you see error 0xC004F074, your internet connection may be unstable.
echo Or the KMS server might be busy. You can try running this script again later.
echo.
pause