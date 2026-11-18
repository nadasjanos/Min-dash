@echo off
setlocal
REM ---------------------------------------------------------------------------
REM inditas.bat
REM EZ A FAJL: dupla kattintassal MINDENT telepit, ami kell, es elinditja
REM a minosegugyi dashboardot. (Szandekosan ekezet nelkul irva, mert a
REM Windows parancsertelmezoje az ekezetes betuket rosszul kezeli.)
REM  1) Ha nincs Python a gepen, feltelepiti (winget-tel, csak a te
REM     felhasznalodnak, rendszergazdai jog nelkul).
REM  2) Ha nincs .venv kornyezet (vagy serult), letrehozza.
REM  3) Feltelepiti / ellenorzi a requirements.txt csomagjait.
REM  4) Letrehozza a data\real mappat, es ha kell, a szintetikus adatot.
REM  5) Elinditja az appot, ami magatol megnyilik a bongeszoben.
REM Leallitas: zard be ezt a fekete ablakot.
REM ---------------------------------------------------------------------------

REM Atlepunk abba a mappaba, ahol ez a fajl van
cd /d "%~dp0"
set "VENV_PY=.venv\Scripts\python.exe"

REM --- 1) Python keresese ---------------------------------------------------
REM Ha a .venv mar mukodik, nem kell kulon Python keresni
if exist "%VENV_PY%" (
    "%VENV_PY%" --version >nul 2>&1
    if not errorlevel 1 goto csomagok
)

set PY=
set PY312="%LOCALAPPDATA%\Programs\Python\Python312\python.exe"

python --version >nul 2>&1
if not errorlevel 1 set PY=python
if defined PY goto venv

py -3 --version >nul 2>&1
if not errorlevel 1 set PY=py -3
if defined PY goto venv

if exist %PY312% set PY=%PY312%
if defined PY goto venv

echo Nincs Python a gepen. Telepites winget-tel, kerlek varj...
winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
REM Az uj Python meg nincs a PATH-on ebben az ablakban, ezert teljes uttal hivjuk
if not exist %PY312% goto nincs_winget
set PY=%PY312%

:venv
REM --- 2) .venv letrehozasa -------------------------------------------------
echo Python kornyezet letrehozasa...
if exist ".venv" rmdir /s /q ".venv"
%PY% -m venv .venv
if errorlevel 1 goto hiba

:csomagok
REM --- 3) Csomagok telepitese / ellenorzese ---------------------------------
echo Csomagok ellenorzese, telepitese - elso alkalommal ez par perc...
"%VENV_PY%" -m pip install -q --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto hiba

REM --- 4) Mappak es szintetikus adat ----------------------------------------
if not exist "data\real" mkdir "data\real"
if exist "data\synthetic\nemmegfelelosegek.csv" goto indit
echo Szintetikus adat eloallitasa...
"%VENV_PY%" generate_data.py
if errorlevel 1 goto hiba

:indit
REM A Streamlit elso inditaskor e-mail cimet kerne. Ezt egy ures beallitassal
REM kihagyjuk (ugyanazt irja, mintha Entert nyomnal a kerdesnel).
if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit"
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    echo [general]> "%USERPROFILE%\.streamlit\credentials.toml"
    echo email = "">> "%USERPROFILE%\.streamlit\credentials.toml"
)

REM --- 5) Inditas -----------------------------------------------------------
echo.
echo A dashboard indul, a bongeszo magatol megnyilik.
echo Leallitashoz zard be ezt az ablakot.
echo.
"%VENV_PY%" -m streamlit run app.py --server.address localhost
goto vege

:nincs_winget
echo.
echo A Python automatikus telepitese nem sikerult.
echo Telepitsd kezzel: https://www.python.org/downloads/
echo Telepiteskor pipald be: Add python.exe to PATH. Utana inditsd ujra ezt a fajlt.
goto vege

:hiba
echo.
echo HIBA tortent. Keszits kepet errol az ablakrol, es kerj segitseget.

:vege
pause
