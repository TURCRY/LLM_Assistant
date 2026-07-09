@echo off
setlocal EnableExtensions EnableDelayedExpansion
:: WAV -> mono 16-bit with visible progress (ffmpeg -stats)

set "SR=16000"
set "FORMAT=WAV"
set "OVERWRITE=1"
set "THREADS=0"
set "RESAMPLE_PRESET=medium"

for %%I in (ffmpeg.exe) do set "FFMPEG=%%~$PATH:I"
if not defined FFMPEG (
  echo [ERROR] ffmpeg not found in PATH. Install ffmpeg and retry.
  exit /b 1
)

set "IN="
if not "%~1"=="" (
  set "IN=%~1"
) else (
  for %%F in ("%~dp0*.wav") do if not defined IN set "IN=%%~fF"
)

if not defined IN (
  echo [ERROR] No .wav provided and none found next to the script.
  exit /b 1
)

if not exist "%IN%" (
  echo [ERROR] Input not found: "%IN%"
  exit /b 1
)

for %%A in ("%IN%") do (
  set "IN_DIR=%%~dpA"
  set "IN_NAME=%%~nxA"
  set "STEM=%%~nA"
)

if /i "%FORMAT%"=="WAV" (
  set "OUT=%IN_DIR%%STEM%_mono16_%SR%Hz.wav"
  set "MUXFMT=-f wav"
  set "AUDIOOPTS=-ac 1 -c:a pcm_s16le"
) else if /i "%FORMAT%"=="FLAC" (
  set "OUT=%IN_DIR%%STEM%_mono16_%SR%Hz.flac"
  set "MUXFMT=-f flac"
  set "AUDIOOPTS=-ac 1 -c:a flac -compression_level 5 -sample_fmt s16"
) else (
  echo [ERROR] Invalid FORMAT: "%FORMAT%". Use WAV or FLAC.
  exit /b 1
)

set "FFTHREADS=-threads %THREADS%"
set "Y="
if "%OVERWRITE%"=="1" ( set "Y=-y" ) else ( set "Y=-n" )
set "TMP=%OUT%.part"

echo ---------------------------------------------------------------
echo Input  : "%IN%"
echo Output : "%OUT%"
echo Format : %FORMAT%  -  SR=%SR% Hz  -  mono 16-bit
echo ffmpeg : "%FFMPEG%"
echo ---------------------------------------------------------------

title Converting: %STEM%
echo Starting conversion...
"%FFMPEG%" -hide_banner -v error -stats -nostdin %FFTHREADS% -i "%IN%" -ar %SR% %AUDIOOPTS% %Y% %MUXFMT% "%TMP%"

if errorlevel 1 (
  echo.
  echo [FAILED] Conversion failed. See errors above.
  if exist "%TMP%" del "%TMP%" >nul 2>&1
  exit /b 1
)

if not exist "%TMP%" (
  echo.
  echo [FAILED] Temporary output was not created.
  exit /b 1
)

if exist "%OUT%" (
  if "%OVERWRITE%"=="1" del "%OUT%" >nul 2>&1
)

move /y "%TMP%" "%OUT%" >nul
if errorlevel 1 (
  echo [ERROR] Could not write output: "%OUT%"
  exit /b 1
)

set "SRC_PATH=%IN%"
set "DST_PATH=%OUT%"

powershell -NoLogo -NoProfile -Command ^
  "$src = Get-Item -LiteralPath $env:SRC_PATH; " ^
  "$dst = Get-Item -LiteralPath $env:DST_PATH; " ^
  "$dst.CreationTime = $src.CreationTime; " ^
  "$dst.LastWriteTime = $src.LastWriteTime"

echo.
echo [OK] Done: "%OUT%"
exit /b 0
