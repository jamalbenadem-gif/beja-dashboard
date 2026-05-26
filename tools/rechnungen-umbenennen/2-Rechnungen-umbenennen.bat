@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo ============================================================
echo   Rechnungen umbenennen
echo ============================================================
echo.

REM --- Python suchen ---
set "PYEXE="
where python >nul 2>&1 && set "PYEXE=python"
if not defined PYEXE (
  where py >nul 2>&1 && set "PYEXE=py"
)
if not defined PYEXE (
  echo [FEHLER] Python nicht gefunden. Bitte zuerst "1-Einrichten.bat" ausfuehren.
  echo.
  pause
  exit /b 1
)

REM --- Ordner ermitteln: per Drag^&Drop oder Eingabe ---
set "ORDNER=%~1"
if not defined ORDNER (
  echo Tipp: Du kannst einen Ordner auch direkt auf diese Datei ziehen.
  echo.
  set /p ORDNER=Pfad zum Rechnungs-Ordner:
)
set ORDNER=!ORDNER:"=!

REM --- Ordner pruefen (robust, auch mit Leerzeichen) ---
pushd "!ORDNER!" 2>nul
if errorlevel 1 (
  echo [FEHLER] Ordner nicht gefunden: !ORDNER!
  echo.
  pause
  exit /b 1
)
popd

echo.
echo === VORSCHAU - es wird noch NICHTS umbenannt ===
echo.
%PYEXE% "%~dp0rechnungen_umbenennen.py" "!ORDNER!"
if errorlevel 1 (
  echo.
  echo [FEHLER] Es ist ein Problem aufgetreten ^(siehe Meldung oben^).
  echo Tipp: Schluessel fehlt? Dann "1-Einrichten.bat" ausfuehren.
  pause
  exit /b 1
)

REM --- Vorschau-Tabelle oeffnen ---
if exist "!ORDNER!\umbenennungen_vorschau.csv" start "" "!ORDNER!\umbenennungen_vorschau.csv"

echo.
echo Bitte die geoeffnete Tabelle (umbenennungen_vorschau.csv) pruefen.
set /p ANTWORT=Jetzt wirklich umbenennen? (j/n):
if /i "!ANTWORT!"=="j" (
  echo.
  echo === UMBENENNEN ===
  echo.
  %PYEXE% "%~dp0rechnungen_umbenennen.py" "!ORDNER!" --apply
) else (
  echo Abgebrochen. Es wurde nichts umbenannt.
)

echo.
pause
