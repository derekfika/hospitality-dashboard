@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo.
echo ============================================================
echo  Commit projects separately and push to Git
echo ============================================================
echo.
echo Separate commits will be created for:
echo   1. Client Booking Platform
echo   2. CPU Dashboard
echo   3. 58VE Dashboard
echo   4. Angel Court Hospitality Dashboard and root scripts
echo.
echo Existing unstaged changes inside those project folders will
echo be included. Nothing outside this workspace will be touched.
echo.
choice /C YN /N /M "Continue? [Y/N]: "
if errorlevel 2 goto :cancelled

set "MESSAGE="
set /P "MESSAGE=Commit description [Update hospitality platforms]: "
if not defined MESSAGE set "MESSAGE=Update hospitality platforms"

call :commit_folder "Client Booking Platform" "client-booking-platform" "%MESSAGE% - booking platform"
if errorlevel 1 goto :failed

call :commit_folder "CPU Dashboard" "cpu-dashboard" "%MESSAGE% - CPU dashboard"
if errorlevel 1 goto :failed

call :commit_folder "58VE Dashboard" "58ve-dashboard" "%MESSAGE% - 58VE dashboard"
if errorlevel 1 goto :failed

echo.
echo ------------------------------------------------------------
echo  Staging Angel Court Hospitality Dashboard
echo ------------------------------------------------------------
git add -- ^
  .claspignore .gitignore ^
  00_config.js 02_Schema.js 03_Utils.js 04_Parser.js ^
  05_GmailScanner.js 06_DataLayer.js 07_Webapp.js ^
  08_DriveHelper.js 09_QuoteEngine.js 10_Calendar.js ^
  11_Triggers.js 12_TestHarness.js appsscript.json ^
  Index.html Styles.html Script.html Icons.html ^
  README.md CHANGELOG.md ^
  push-apps-script.bat push-git-projects.bat
if errorlevel 1 goto :failed

call :commit_staged "%MESSAGE% - Angel Court dashboard"
if errorlevel 1 goto :failed

for /F "delims=" %%B in ('git branch --show-current') do set "BRANCH=%%B"
if not defined BRANCH (
  echo ERROR: Could not determine the current Git branch.
  goto :failed
)

echo.
echo ------------------------------------------------------------
echo  Pushing branch %BRANCH%
echo ------------------------------------------------------------
git push origin "%BRANCH%"
if errorlevel 1 goto :failed

echo.
echo ============================================================
echo  Project commits and Git push completed successfully.
echo ============================================================
echo.
git status --short
echo.
pause
exit /b 0

:commit_folder
echo.
echo ------------------------------------------------------------
echo  Staging %~1
echo ------------------------------------------------------------
git add -A -- "%~2"
if errorlevel 1 exit /b 1
call :commit_staged "%~3"
exit /b %ERRORLEVEL%

:commit_staged
git diff --cached --quiet
if not errorlevel 1 (
  echo No changes to commit for this project.
  exit /b 0
)

git commit -m "%~1"
if errorlevel 1 exit /b 1
exit /b 0

:failed
echo.
echo ============================================================
echo  Git operation stopped because a step failed.
echo  Review the message above before running the file again.
echo ============================================================
echo.
git status --short
echo.
pause
exit /b 1

:cancelled
echo.
echo Git operation cancelled.
exit /b 0
