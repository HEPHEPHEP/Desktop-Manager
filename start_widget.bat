@echo off
:: Desktop Folder Widget Starter v3.0
:: ===================================

title Desktop Folder Widget

cd /d "%~dp0"

:: Prüfen ob Python installiert ist
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [FEHLER] Python wurde nicht gefunden!
    echo.
    echo Bitte installieren Sie Python von https://python.org
    echo.
    pause
    exit /b 1
)

:: Widget starten mit sichtbarer Konsole für Debug
echo ========================================
echo   Desktop Folder Widget v3.0
echo ========================================
echo.
echo Bedienung:
echo   - Dateien auf die Kachel ziehen = Hinzufuegen
echo   - Icons aus der Kachel ziehen   = Wiederherstellen  
echo   - Rechtsklick                   = Kontextmenue
echo.
echo WICHTIG: Beim Beenden werden alle Icons wiederhergestellt!
echo.
echo ========================================
echo.

python desktop_folder_widget_v3.py

echo.
echo ========================================
echo   Widget beendet
echo ========================================
echo.
echo Falls Desktop-Icons noch versteckt sind:
echo   restore_desktop_icons.bat ausfuehren
echo.
pause
