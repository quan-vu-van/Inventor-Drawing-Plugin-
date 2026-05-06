@echo off
setlocal enabledelayedexpansion

:: 1. Define paths
set "SOURCE_DIR=C:\CustomTools\Inventor"
set "INVENTOR_USER_ADDIN=%APPDATA%\Autodesk\Inventor 2023\Addins"

echo =====================================================
echo    ACTIVATING MCG_ ADD-INS FOR INVENTOR 2023
echo =====================================================

:: 2. Check source directory
if not exist "%SOURCE_DIR%" (
    echo [ERROR] Source directory not found: %SOURCE_DIR%
    pause
    exit /b
)

:: 3. Create Inventor Addins folder if it doesn't exist
if not exist "%INVENTOR_USER_ADDIN%" (
    echo [+] Creating Addins directory...
    mkdir "%INVENTOR_USER_ADDIN%"
)

:: 4. Loop through MCG_ folders
echo [+] Scanning for MCG_ Add-ins...
set /a count=0
set /a success=0

for /D %%D in ("%SOURCE_DIR%\MCG_*") do (
    set "FOLDER_NAME=%%~nxD"
    set "ADDIN_FILE=%%D\!FOLDER_NAME!.addin"
    set "DLL_FILE=%%D\!FOLDER_NAME!.dll"
    
    echo.
    echo -- Processing: !FOLDER_NAME!
    
    :: Check if both .addin and .dll exist
    if exist "!ADDIN_FILE!" (
        if exist "!DLL_FILE!" (
            set /a count+=1
            
            :: Create target folder
            set "TARGET_FOLDER=%INVENTOR_USER_ADDIN%\!FOLDER_NAME!"
            if not exist "!TARGET_FOLDER!" (
                mkdir "!TARGET_FOLDER!"
                echo    [+] Created folder: !FOLDER_NAME!
            ) else (
                echo    [i] Folder exists, overwriting files...
            )
            
            :: Copy .addin file
            copy /Y "!ADDIN_FILE!" "!TARGET_FOLDER!\" >nul 2>&1
            if !errorlevel! equ 0 (
                echo    [OK] Copied: !FOLDER_NAME!.addin
            ) else (
                echo    [!] Failed to copy: !FOLDER_NAME!.addin
            )
            
            :: Copy .dll file
            copy /Y "!DLL_FILE!" "!TARGET_FOLDER!\" >nul 2>&1
            if !errorlevel! equ 0 (
                echo    [OK] Copied: !FOLDER_NAME!.dll
                set /a success+=1
            ) else (
                echo    [!] Failed to copy: !FOLDER_NAME!.dll
            )
        ) else (
            echo    [!] Missing DLL: !FOLDER_NAME!.dll
        )
    ) else (
        echo    [!] Missing ADDIN: !FOLDER_NAME!.addin
    )
)

echo.
echo =====================================================
if %count% gtr 0 (
    echo [SUMMARY] Found %count% MCG_ Add-in(s)
    echo           Successfully activated: %success%
    echo.
    echo Target folder: %INVENTOR_USER_ADDIN%
) else (
    echo [!] No MCG_ Add-ins found in %SOURCE_DIR%
)
echo =====================================================
echo.
echo Please restart Inventor 2023 to apply changes.
echo.
pause