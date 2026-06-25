@echo off
REM Run this on a Windows machine that HAS Python installed.
REM It produces sync_agent.exe inside the "dist" folder, which you then
REM copy to the factory PC alongside config.json. The factory PC needs
REM NO Python installation at all.

pip install -r requirements.txt
pyinstaller --onefile --name sync_agent --console sync_agent.py

echo.
echo ============================================
echo Build complete.
echo Find sync_agent.exe inside the "dist" folder.
echo Copy sync_agent.exe + config.json together
echo into one folder on the factory PC.
echo ============================================
pause
