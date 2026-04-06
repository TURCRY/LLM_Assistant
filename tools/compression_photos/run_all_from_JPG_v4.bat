@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM =========================================================
REM PARAMETRES
REM =========================================================
set "AFFAIRE=%~1"
if "%AFFAIRE%"=="" (
    echo ERREUR: id_affaire requis. Exemple : 2026-J58
    exit /b 1
)

set "ROOT_PC1=\\192.168.1.20\Affaires"
set "ROOT_PC2=\\192.168.0.155\Affaires"

set "PY_GEN=C:\DevTools\Compression photo\gen_tout_depuis_jpg.py"
set "PY_POST=C:\DevTools\Compression photo\batch_photos_post_sync.py"

REM =========================================================
REM CONTROLES DOSSIER COURANT
REM Doit etre lance depuis ...\<id_captation>\photos\JPG
REM =========================================================
echo SCRIPT EXECUTE : %~f0
echo CWD           : %CD%

for %%I in ("%CD%") do set "CURR_NAME=%%~nxI"
for %%I in ("%CD%\..") do (
    set "SRC_PHOTOS=%%~fI"
    set "PHOTOS_DIR_NAME=%%~nxI"
)
for %%I in ("%SRC_PHOTOS%\..") do (
    set "CAPTATION_DIR=%%~fI"
    set "CAPTATION=%%~nxI"
)

if /I not "%CURR_NAME%"=="JPG" (
    echo ERREUR: le script doit etre lance depuis ...\photos\JPG
    exit /b 1
)

if not "%PHOTOS_DIR_NAME%"=="photos" (
    echo ERREUR: le dossier parent de JPG doit s'appeler exactement "photos" en minuscules.
    exit /b 1
)

if not exist "%CAPTATION_DIR%\audio" (
    echo ERREUR: dossier "audio" introuvable sous %CAPTATION_DIR%
    exit /b 1
)

set "SRC_AUDIO=%CAPTATION_DIR%\audio"

if "%CAPTATION%"=="" (
    echo ERREUR: impossible de determiner l'id_captation
    exit /b 1
)

REM =========================================================
REM GENERATION PHOTOS
REM =========================================================
echo [1/8] Generation de photos.csv et des JPG reduits...

python "%PY_GEN%"
if errorlevel 1 (
    echo ERREUR: gen_tout_depuis_jpg.py a echoue
    exit /b 1
)

set "PHOTOS_CSV=%SRC_PHOTOS%\photos.csv"
set "PHOTOS_BATCH_CSV=%SRC_PHOTOS%\photos_batch.csv"
set "PHOTOS_XLS=%SRC_PHOTOS%\photos.xls"

if not exist "%PHOTOS_CSV%" (
    echo ERREUR: photos.csv non genere dans %SRC_PHOTOS%
    exit /b 1
)

REM =========================================================
REM DETECTION FICHIERS AUDIO / TRANSCRIPTIONS
REM =========================================================
set "WAV_MONO16="
set "WAV_SOURCE="
set "TRANSCRIPT_CSV="
set "CTX_GENERAL="

for %%F in ("%SRC_AUDIO%\*_mono16_16000Hz.wav") do (
    if exist "%%~fF" (
        set "WAV_MONO16=%%~fF"
        goto :wav16_found
    )
)
:wav16_found

for %%F in ("%SRC_AUDIO%\*.wav") do (
    if exist "%%~fF" if /I not "%%~fF"=="%WAV_MONO16%" (
        set "WAV_SOURCE=%%~fF"
        goto :wavsrc_found
    )
)
:wavsrc_found

for %%F in ("%SRC_AUDIO%\*.csv") do (
    echo %%~nxF | find /I "(photo).csv" >nul
    if errorlevel 1 (
        echo %%~nxF | find /I "photos.csv" >nul
        if errorlevel 1 (
            echo %%~nxF | find /I "photos_batch.csv" >nul
            if errorlevel 1 (
                set "TRANSCRIPT_CSV=%%~fF"
                goto :transcript_found
            )
        )
    )
)
:transcript_found

for %%F in ("%SRC_AUDIO%\contexte_general*.json") do (
    if exist "%%~fF" (
        set "CTX_GENERAL=%%~fF"
        goto :ctx_found
    )
)
:ctx_found

REM =========================================================
REM AFFICHAGE CONTROLE
REM =========================================================
echo ----------------------------------------
echo AFFAIRE        = %AFFAIRE%
echo CAPTATION      = %CAPTATION%
echo ROOT_PC1       = %ROOT_PC1%
echo ROOT_PC2       = %ROOT_PC2%
echo SRC_PHOTOS     = %SRC_PHOTOS%
echo SRC_AUDIO      = %SRC_AUDIO%
echo PHOTOS_CSV     = %PHOTOS_CSV%
echo TRANSCRIPT_CSV = %TRANSCRIPT_CSV%
echo WAV_MONO16     = %WAV_MONO16%
echo WAV_SOURCE     = %WAV_SOURCE%
echo CTX_GENERAL    = %CTX_GENERAL%
echo ----------------------------------------

REM =========================================================
REM COPIE VERS NAS
REM =========================================================
echo [2/8] Copie vers NAS ASUSTOR...
call :copy_to_target "%ROOT_PC1%"
if errorlevel 1 exit /b 1

REM =========================================================
REM COPIE VERS PC FIXE
REM =========================================================
echo [3/8] Copie vers PC fixe Windows...
call :copy_to_target "%ROOT_PC2%"
if errorlevel 1 exit /b 1

REM =========================================================
REM MISE A JOUR DU CSV LOCAL
REM =========================================================
echo [4/8] Mise a jour photos.csv ^(champs pcfixe^)...

python "%PY_POST%" ^
  --photos_csv "%PHOTOS_CSV%" ^
  --id_affaire "%AFFAIRE%" ^
  --id_captation "%CAPTATION%" ^
  --root_pcfixe "%ROOT_PC2%"

if errorlevel 1 (
    echo ERREUR: batch_photos_post_sync.py a echoue
    exit /b 1
)

echo OK - batch_photos_post_sync.py termine

REM =========================================================
REM RECOPIE DES CSV/XLS MIS A JOUR VERS LES 2 CIBLES
REM =========================================================
echo [5/8] Recopie des CSV/XLS mis a jour vers NAS...
call :copy_csv_to_target "%ROOT_PC1%"
if errorlevel 1 exit /b 1

echo [6/8] Recopie des CSV/XLS mis a jour vers PC fixe...
call :copy_csv_to_target "%ROOT_PC2%"
if errorlevel 1 exit /b 1

echo [7/8] Termine avec succes.
exit /b 0


REM =========================================================
REM FONCTION : COPIE COMPLETE VERS UNE CIBLE
REM =========================================================
:copy_to_target
set "ROOT=%~1"

set "DST_CAPT=%ROOT%\%AFFAIRE%\AE_Expert_captations\%CAPTATION%"
set "DST_PHOTOS=%DST_CAPT%\photos"
set "DST_AUDIO=%DST_CAPT%\audio"
set "DST_ASR=%ROOT%\%AFFAIRE%\AF_Expert_ASR\transcriptions\%CAPTATION%"

echo ----------------------------------------
echo COPIE VERS %ROOT%
echo DST_PHOTOS = %DST_PHOTOS%
echo DST_AUDIO  = %DST_AUDIO%
echo DST_ASR    = %DST_ASR%
echo ----------------------------------------

mkdir "%DST_PHOTOS%" 2>nul
mkdir "%DST_AUDIO%" 2>nul
mkdir "%DST_ASR%" 2>nul

REM 1) photos + sous-dossiers
robocopy "%SRC_PHOTOS%" "%DST_PHOTOS%" /E /R:2 /W:1 /Z /FFT /XA:SH /NP
set "RC=%ERRORLEVEL%"
if %RC% GEQ 8 (
    echo ERREUR: echec copie photos vers %DST_PHOTOS%
    exit /b 1
)

REM 2) wav vers AE_Expert_captations\...\audio
robocopy "%SRC_AUDIO%" "%DST_AUDIO%" *.wav /R:2 /W:1 /Z /FFT /XA:SH /NP
set "RC=%ERRORLEVEL%"
if %RC% GEQ 8 (
    echo ERREUR: echec copie WAV vers %DST_AUDIO%
    exit /b 1
)

REM 2b) dossiers .wav.part vers AE_Expert_captations\...\audio
for /d %%D in ("%SRC_AUDIO%\*.wav.part") do (
    robocopy "%%~fD" "%DST_AUDIO%\%%~nxD" /E /R:2 /W:1 /Z /FFT /XA:SH /NP
    set "RC=!ERRORLEVEL!"
    if !RC! GEQ 8 (
        echo ERREUR: echec copie dossier %%~nxD vers %DST_AUDIO%
        exit /b 1
    )
)

REM 3) hors wav vers AF_Expert_ASR\transcriptions\...
robocopy "%SRC_AUDIO%" "%DST_ASR%" *.csv *.srt *.vtt *.json *.txt *.xlsx *.xls /R:2 /W:1 /Z /FFT /XA:SH /NP
set "RC=%ERRORLEVEL%"
if %RC% GEQ 8 (
    echo ERREUR: echec copie transcriptions vers %DST_ASR%
    exit /b 1
)

echo OK - copie complete vers %ROOT%
exit /b 0


REM =========================================================
REM FONCTION : RECOPIE CSV/XLS MIS A JOUR
REM =========================================================
:copy_csv_to_target
set "ROOT=%~1"
set "DST_PHOTOS=%ROOT%\%AFFAIRE%\AE_Expert_captations\%CAPTATION%\photos"

robocopy "%SRC_PHOTOS%" "%DST_PHOTOS%" photos.csv photos_batch.csv photos.xls /R:2 /W:1 /Z /FFT /XA:SH /NP /NFL /NDL
set "RC=%ERRORLEVEL%"
if %RC% GEQ 8 (
    echo ERREUR: echec recopie CSV/XLS vers %DST_PHOTOS%
    exit /b 1
)

echo OK - CSV/XLS recopie(s) vers %ROOT%
exit /b 0