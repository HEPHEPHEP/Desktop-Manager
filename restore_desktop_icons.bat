@echo off
:: Stellt alle versteckten Desktop-Icons wieder her
:: ================================================

title Desktop Icons Wiederherstellen

echo.
echo ========================================
echo   Desktop Icons Wiederherstellen
echo ========================================
echo.
echo Dieses Skript macht alle versteckten Dateien
echo auf dem Desktop wieder sichtbar.
echo.

:: Desktop-Pfad ermitteln
set DESKTOP=%USERPROFILE%\Desktop

echo Desktop-Pfad: %DESKTOP%
echo.
echo Stelle versteckte Dateien wieder her...
echo.

:: Alle versteckten Dateien auf dem Desktop sichtbar machen
attrib -H "%DESKTOP%\*.*" /S

echo.
echo Fertig! Alle Desktop-Dateien sollten jetzt sichtbar sein.
echo.
echo Falls Dateien immer noch fehlen, druecken Sie F5 auf dem Desktop.
echo.
pause
