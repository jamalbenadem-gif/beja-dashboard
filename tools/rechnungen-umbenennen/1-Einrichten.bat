@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo ============================================================
echo   Einrichtung fuer "Rechnungen umbenennen"  (einmalig)
echo ============================================================
echo.

REM --- Python suchen (python oder py) ---
set "PYEXE="
where python >nul 2>&1 && set "PYEXE=python"
if not defined PYEXE (
  where py >nul 2>&1 && set "PYEXE=py"
)
if not defined PYEXE (
  echo [FEHLER] Python wurde nicht gefunden.
  echo Bitte zuerst Python 3.10+ installieren:
  echo    https://www.python.org/downloads/windows/
  echo Beim Setup unbedingt "Add python.exe to PATH" anhaken,
  echo danach diese Datei erneut starten.
  echo.
  pause
  exit /b 1
)
echo [OK] Python gefunden (!PYEXE!).
echo.

REM --- Pakete installieren ---
echo Installiere "anthropic" ...
%PYEXE% -m pip install -U anthropic
if errorlevel 1 (
  echo [FEHLER] Installation von "anthropic" fehlgeschlagen. Internet pruefen.
  pause
  exit /b 1
)
echo Installiere "pillow" (fuer TIFF/BMP, optional) ...
%PYEXE% -m pip install -U pillow
echo [OK] Pakete installiert.
echo.

REM --- API-Schluessel ---
if defined ANTHROPIC_API_KEY goto KEYDONE
echo Bitte den Anthropic API-Schluessel eingeben.
echo Er beginnt mit sk-ant- und stammt von https://console.anthropic.com
set /p KEY=Schluessel:
if not defined KEY goto KEYSKIP
setx ANTHROPIC_API_KEY "!KEY!" >nul
echo [OK] Schluessel gespeichert.
echo HINWEIS: Dieses Fenster bitte schliessen. Der Schluessel ist ab dem
echo naechsten Start aktiv. Falls er nicht erkannt wird, den PC einmal neu starten.
goto KEYDONE
:KEYSKIP
echo [Uebersprungen] Kein Schluessel eingegeben - das kannst du spaeter nachholen.
:KEYDONE

echo.
echo Fertig. Du kannst jetzt "2-Rechnungen-umbenennen.bat" verwenden.
echo.
pause
