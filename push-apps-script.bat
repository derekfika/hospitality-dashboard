@echo off
setlocal
cd /d "%~dp0"

rem Named clasp users. Authorise each once with:
rem   npx.cmd --yes @google/clasp login --user hospitality
rem   npx.cmd --yes @google/clasp login --user cpu
set "HOSPITALITY_CLASP_USER=hospitality"
set "CPU_CLASP_USER=cpu"

echo.
echo ============================================================
echo  Push all Hospitality Apps Script projects
echo ============================================================
echo.
echo This will push:
echo   1. Angel Court Hospitality Dashboard
echo   2. Client Booking Platform
echo   3. CPU Dashboard
echo   4. 58VE Dashboard
echo.
echo Accounts:
echo   Hospitality projects: %HOSPITALITY_CLASP_USER%
echo   CPU Dashboard:        %CPU_CLASP_USER%
echo.
choice /C YN /N /M "Continue? [Y/N]: "
if errorlevel 2 goto :cancelled

call :push_project "Angel Court Hospitality Dashboard" "%~dp0" "%HOSPITALITY_CLASP_USER%"
if errorlevel 1 goto :failed

call :push_project "Client Booking Platform" "%~dp0client-booking-platform" "%HOSPITALITY_CLASP_USER%"
if errorlevel 1 goto :failed

call :push_project "CPU Dashboard" "%~dp0cpu-dashboard" "%CPU_CLASP_USER%"
if errorlevel 1 goto :failed

call :push_project "58VE Dashboard" "%~dp058ve-dashboard" "%HOSPITALITY_CLASP_USER%"
if errorlevel 1 goto :failed

echo.
echo ============================================================
echo  All Apps Script projects pushed successfully.
echo ============================================================
echo.
pause
exit /b 0

:push_project
echo.
echo ------------------------------------------------------------
echo  Pushing %~1
echo  Clasp user: %~3
echo ------------------------------------------------------------

if not exist "%~2\.clasp.json" (
  echo ERROR: Missing .clasp.json in "%~2".
  exit /b 1
)

pushd "%~2"

where clasp.cmd >nul 2>nul
if not errorlevel 1 (
  call clasp.cmd --user "%~3" push
) else (
  where npx.cmd >nul 2>nul
  if errorlevel 1 (
    echo ERROR: Neither clasp nor npx was found.
    popd
    exit /b 1
  )
  call npx.cmd --yes @google/clasp --user "%~3" push
)

set "PUSH_RESULT=%ERRORLEVEL%"
popd

if not "%PUSH_RESULT%"=="0" (
  echo ERROR: %~1 failed to push.
  exit /b %PUSH_RESULT%
)

echo SUCCESS: %~1
exit /b 0

:failed
echo.
echo ============================================================
echo  Push stopped because a project failed.
echo  Later projects were not pushed.
echo ============================================================
echo.
pause
exit /b 1

:cancelled
echo.
echo Push cancelled.
exit /b 0
