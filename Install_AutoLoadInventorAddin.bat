@echo off
setlocal enabledelayedexpansion

:: ============================================================
:: MCG Inventor Add-In Installer
:: Scans all .addin files in SOURCE_DIR and copies them to
:: Inventor's user Addins folder. DLLs stay in SOURCE_DIR
:: (the .addin files reference DLLs by absolute path).
::
:: Source layout (flat):
::   C:\CustomTools\Inventor\
::   ├── MCG_InventorCreateDummyDetailSection.dll
::   ├── MCG_InventorCreateDummyDetailSection.addin
::   ├── Symbol Sections.idw
::   ├── ... (other addins' .dll + .addin)
::   └── Install_AutoLoadInventorAddin.bat
::
:: Target: %APPDATA%\Autodesk\Inventor 2023\Addins\
:: ============================================================

set "SOURCE_DIR=C:\CustomTools\Inventor"
set "INVENTOR_USER_ADDIN=%APPDATA%\Autodesk\Inventor 2023\Addins"

echo =====================================================
echo    MCG Inventor Add-In Installer
echo =====================================================

if not exist "%SOURCE_DIR%" (
    echo [ERROR] Source directory not found: %SOURCE_DIR%
    pause
    exit /b
)

if not exist "%INVENTOR_USER_ADDIN%" (
    echo [+] Creating Addins directory...
    mkdir "%INVENTOR_USER_ADDIN%"
)

echo [+] Scanning %SOURCE_DIR% for .addin files...
set /a count=0

for %%F in ("%SOURCE_DIR%\*.addin") do (
    echo -- Copying: %%~nxF
    copy /Y "%%F" "%INVENTOR_USER_ADDIN%\" >nul
    if !errorlevel! equ 0 (
        set /a count+=1
    ) else (
        echo [!] Error copying file: %%~nxF
    )
)

echo -----------------------------------------------------
if %count% gtr 0 (
    echo [OK] Success! %count% Add-in(s) activated.
    echo Target folder: %INVENTOR_USER_ADDIN%
) else (
    echo [!] No .addin files found in %SOURCE_DIR%
)

echo -----------------------------------------------------
echo Please restart Inventor 2023 to apply changes.
pause
