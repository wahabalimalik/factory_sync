@echo off
REM Run this on a Windows machine that HAS Python installed.
REM IMPORTANT: extract the zip FULLY first (right-click -> Extract All)
REM and run this .bat from the real extracted folder - not from inside
REM the zip viewer, or the output will vanish when you close it.

echo Running from: %cd%
echo.

echo Step 1/3: Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ============================================
    echo  FAILED at "pip install". See the error
    echo  above. Common cause: no internet access, or
    echo  pip/python not on PATH.
    echo ============================================
    pause
    exit /b 1
)

echo.
echo Step 2/3: Building the exe with PyInstaller...
pyinstaller --onefile --name sync_agent --console sync_agent.py
if errorlevel 1 (
    echo.
    echo ============================================
    echo  FAILED at "pyinstaller". See the error
    echo  above.
    echo ============================================
    pause
    exit /b 1
)

echo.
echo Step 3/3: Verifying the exe actually exists...
if not exist "dist\sync_agent.exe" (
    echo.
    echo ============================================
    echo  PyInstaller reported success but
    echo  dist\sync_agent.exe is MISSING.
    echo  This almost always means your antivirus
    echo  deleted it right after creation.
    echo  Check: Windows Security -^> Protection history.
    echo  Add this folder to Defender exclusions, then
    echo  run this script again.
    echo ============================================
    pause
    exit /b 1
)

echo.
echo ============================================
echo  SUCCESS - verified the file exists at:
echo  %cd%\dist\sync_agent.exe
echo.
echo  Copy sync_agent.exe + config.json together
echo  into one folder on the factory PC.
echo ============================================
pause
