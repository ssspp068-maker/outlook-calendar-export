@echo off
chcp 65001 >nul
setlocal EnableExtensions
set "SRC=%~dp0"
set "OTM_SRC=%SRC%VBAProject.OTM"
set "UI_SRC=%SRC%olkexplorer.officeUI"
set "BAS=%SRC%ExportCalendarMeetings.bas"
set "VBS=%SRC%import-macro.vbs"
set "OTM_DIR=%UserProfile%\AppData\Roaming\Microsoft\Outlook"
set "UI_DIR=%UserProfile%\AppData\Local\Microsoft\Office"
set "IMPORT_OK=0"

if not exist "%UI_SRC%" (
  echo [ERROR] Missing olkexplorer.officeUI
  pause
  exit /b 1
)

echo Closing Office apps that lock Outlook VBA / OTM ...
taskkill /f /im OUTLOOK.EXE >nul 2>&1
taskkill /f /im EXCEL.EXE >nul 2>&1
taskkill /f /im WINWORD.EXE >nul 2>&1
taskkill /f /im POWERPNT.EXE >nul 2>&1
taskkill /f /im ONENOTE.EXE >nul 2>&1
taskkill /f /im ONENOTEM.EXE >nul 2>&1
taskkill /f /im MSACCESS.EXE >nul 2>&1
taskkill /f /im VISIO.EXE >nul 2>&1
taskkill /f /im WINPROJ.EXE >nul 2>&1
taskkill /f /im LYNC.EXE >nul 2>&1
rem Teams often unrelated to OTM lock; kill only classic Lync/Skype for Business above.
timeout /t 5 /nobreak >nul

if not exist "%OTM_DIR%" mkdir "%OTM_DIR%"
if not exist "%UI_DIR%" mkdir "%UI_DIR%"

IF EXIST "%UI_DIR%\olkexplorer1.officeUI" (
  ECHO Backup olkexplorer1.officeUI already exists
) ELSE (
  IF EXIST "%UI_DIR%\olkexplorer.officeUI" ren "%UI_DIR%\olkexplorer.officeUI" olkexplorer1.officeUI
)

echo Copying Calendar ribbon button ...
xcopy /y "%UI_SRC%" "%UI_DIR%\" /I >nul

rem --- Preferred path: prebuilt VBAProject.OTM (no AccessVBOM / VBE needed) ---
if exist "%OTM_SRC%" (
  echo Found VBAProject.OTM - installing by file copy ^(no Trust Center^) ...
  IF EXIST "%OTM_DIR%\VBAProject1.OTM" (
    ECHO Backup VBAProject1.OTM already exists
  ) ELSE (
    IF EXIST "%OTM_DIR%\VBAProject.OTM" ren "%OTM_DIR%\VBAProject.OTM" VBAProject1.OTM
  )
  xcopy /y "%OTM_SRC%" "%OTM_DIR%\" /I >nul
  if errorlevel 1 (
    echo [ERROR] Failed to copy VBAProject.OTM
    pause
    exit /b 1
  )
  echo VBAProject.OTM installed.
  goto done_ok
)

rem --- Fallback: force AccessVBOM in registry, then import .bas via VBE ---
echo No VBAProject.OTM in package - enabling VBA project access via registry ...
call :EnableAccessVBOM 16.0
call :EnableAccessVBOM 15.0

echo Verifying AccessVBOM ...
set "VBOM_OK=0"
reg query "HKCU\Software\Microsoft\Office\16.0\Outlook\Security" /v AccessVBOM 2>nul | findstr /i "0x1" >nul && set "VBOM_OK=1"
reg query "HKCU\Software\Microsoft\Office\16.0\Common\Security" /v AccessVBOM 2>nul | findstr /i "0x1" >nul && set "VBOM_OK=1"
reg query "HKCU\Software\Microsoft\Office\15.0\Outlook\Security" /v AccessVBOM 2>nul | findstr /i "0x1" >nul && set "VBOM_OK=1"
if "%VBOM_OK%"=="0" (
  echo [WARN] Could not confirm AccessVBOM=1 via reg query. Continuing anyway...
) else (
  echo AccessVBOM=1 confirmed.
)

if not exist "%SRC%ExportCalendarMeetings.bas.p0" goto no_parts
if not exist "%SRC%ExportCalendarMeetings.bas.p1" goto no_parts
if not exist "%SRC%ExportCalendarMeetings.bas.p2" goto no_parts
if not exist "%SRC%ExportCalendarMeetings.bas.p3" goto no_parts
if not exist "%SRC%ExportCalendarMeetings.bas.p4" goto no_parts
if not exist "%SRC%ExportCalendarMeetings.bas.p5" goto no_parts
if not exist "%SRC%ExportCalendarMeetings.bas.p6" goto no_parts
if not exist "%SRC%ExportCalendarMeetings.bas.p7" goto no_parts

echo Assembling ExportCalendarMeetings.bas from p0-p7 ...
copy /b "%SRC%ExportCalendarMeetings.bas.p0"+"%SRC%ExportCalendarMeetings.bas.p1"+"%SRC%ExportCalendarMeetings.bas.p2"+"%SRC%ExportCalendarMeetings.bas.p3"+"%SRC%ExportCalendarMeetings.bas.p4"+"%SRC%ExportCalendarMeetings.bas.p5"+"%SRC%ExportCalendarMeetings.bas.p6"+"%SRC%ExportCalendarMeetings.bas.p7" "%BAS%" >nul
goto have_bas

:no_parts
if not exist "%BAS%" (
  echo [ERROR] Missing ExportCalendarMeetings.bas.p0-p7
  pause
  exit /b 1
)

:have_bas
if not exist "%BAS%" (
  echo [ERROR] Missing ExportCalendarMeetings.bas
  pause
  exit /b 1
)

if not exist "%VBS%" (
  echo [ERROR] Missing import-macro.vbs
  goto last_resort
)

echo Importing macro into Outlook VBA ...
cscript //nologo "%VBS%" "%BAS%"
if errorlevel 1 (
  echo.
  echo Import failed after AccessVBOM registry enable.
  goto last_resort
)
set "IMPORT_OK=1"
goto done_ok

:last_resort
echo.
echo ========================================
echo LAST RESORT ^(only if install failed^):
echo   1. File -^> Options -^> Trust Center -^> Trust Center Settings
echo   2. Macro Settings -^> enable "Trust access to the VBA project object model"
echo   3. Close Outlook fully, run install.bat again
echo ========================================
echo.
echo Or manually: Alt+F11 -^> remove ExportCalendarMeetings -^> Import File -^> ExportCalendarMeetings.bas
echo.
pause
exit /b 1

:done_ok
echo.
echo ========================================
echo OK. Open Outlook -^> Calendar -^> "Выгрузить встречи"
echo Or always: Alt+F8 -^> ExportManagerCalendarMeetings
echo NEVER run InstallCalendarExportButton
echo ========================================
pause
exit /b 0

:EnableAccessVBOM
rem %~1 = Office version e.g. 16.0
reg add "HKCU\Software\Microsoft\Office\%~1\Outlook\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\%~1\Common\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Policies\Microsoft\Office\%~1\Outlook\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Policies\Microsoft\Office\%~1\Common\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul 2>&1
exit /b 0
