@echo off
chcp 65001 >nul
setlocal
set "SRC=%~dp0"
set "UI_DIR=%UserProfile%\AppData\Local\Microsoft\Office"
set "BAS=%SRC%ExportCalendarMeetings.bas"
set "VBS=%SRC%import-macro.vbs"
set "UI=%SRC%olkexplorer.officeUI"
set "B64=%SRC%ExportCalendarMeetings.bas.gz.b64"

if not exist "%UI%" (
  echo [ERROR] Missing olkexplorer.officeUI
  pause
  exit /b 1
)

REM 1) Prefer ASCII ChrW fragments p0-p7 (survive GitHub UTF-8).
if exist "%SRC%ExportCalendarMeetings.bas.p0" if exist "%SRC%ExportCalendarMeetings.bas.p7" (
  echo Assembling ExportCalendarMeetings.bas from p0-p7 ...
  copy /b "%SRC%ExportCalendarMeetings.bas.p0"+"%SRC%ExportCalendarMeetings.bas.p1"+"%SRC%ExportCalendarMeetings.bas.p2"+"%SRC%ExportCalendarMeetings.bas.p3"+"%SRC%ExportCalendarMeetings.bas.p4"+"%SRC%ExportCalendarMeetings.bas.p5"+"%SRC%ExportCalendarMeetings.bas.p6"+"%SRC%ExportCalendarMeetings.bas.p7" "%BAS%" >nul
  goto have_bas
)

REM 2) Else expand from gz.b64 (join gz parts only if pack missing).
if not exist "%B64%" (
  if exist "%SRC%ExportCalendarMeetings.bas.gz.b64.part0" if exist "%SRC%ExportCalendarMeetings.bas.gz.b64.part1" if exist "%SRC%ExportCalendarMeetings.bas.gz.b64.part2" if exist "%SRC%ExportCalendarMeetings.bas.gz.b64.part3" (
    echo Joining bas.gz.b64 parts 0-3 ...
    copy /b "%SRC%ExportCalendarMeetings.bas.gz.b64.part0"+"%SRC%ExportCalendarMeetings.bas.gz.b64.part1"+"%SRC%ExportCalendarMeetings.bas.gz.b64.part2"+"%SRC%ExportCalendarMeetings.bas.gz.b64.part3" "%B64%" >nul
  ) else if exist "%SRC%ExportCalendarMeetings.bas.gz.b64.part0" if exist "%SRC%ExportCalendarMeetings.bas.gz.b64.part1" if exist "%SRC%ExportCalendarMeetings.bas.gz.b64.part2" (
    echo Joining bas.gz.b64 parts 0-2 ...
    copy /b "%SRC%ExportCalendarMeetings.bas.gz.b64.part0"+"%SRC%ExportCalendarMeetings.bas.gz.b64.part1"+"%SRC%ExportCalendarMeetings.bas.gz.b64.part2" "%B64%" >nul
  )
)

if exist "%B64%" (
  echo Expanding ExportCalendarMeetings.bas from gz.b64 ...
  powershell -NoProfile -Command "$b=[Convert]::FromBase64String(((Get-Content -Raw '%B64%') -replace '\s','')); $ms=New-Object IO.MemoryStream(,$b); $gz=New-Object IO.Compression.GzipStream($ms,[IO.Compression.CompressionMode]::Decompress); $out=New-Object IO.FileStream('%BAS%',[IO.FileMode]::Create); $gz.CopyTo($out); $out.Close(); $gz.Close(); $ms.Close()"
  if errorlevel 1 (
    echo [ERROR] Failed to expand bas.gz.b64
    pause
    exit /b 1
  )
)

:have_bas
if not exist "%BAS%" (
  echo [ERROR] Missing ExportCalendarMeetings.bas
  pause
  exit /b 1
)

echo If VBA break mode: Reset (square / Run - Reset) first, then close Outlook.
echo Closing Outlook...
taskkill /f /im OUTLOOK.EXE >nul 2>&1
timeout /t 2 /nobreak >nul

reg add "HKCU\Software\Microsoft\Office\16.0\Outlook\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\Common\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul

if not exist "%UI_DIR%" mkdir "%UI_DIR%"

IF EXIST "%UI_DIR%\olkexplorer1.officeUI" (
  ECHO Backup olkexplorer1.officeUI already exists
) ELSE (
  IF EXIST "%UI_DIR%\olkexplorer.officeUI" ren "%UI_DIR%\olkexplorer.officeUI" olkexplorer1.officeUI
)

xcopy /y "%UI%" "%UI_DIR%\" /I
echo Ribbon button copied to Local\Microsoft\Office

echo Importing macro (deletes old ExportCalendarMeetings module first) ...
if exist "%VBS%" (
  cscript //nologo "%VBS%" "%BAS%"
  if errorlevel 1 (
    echo.
    echo Auto-import failed. Manual once:
    echo   1. Alt+F11 -^> delete module ExportCalendarMeetings (Remove -^> No)
    echo   2. File -^> Import File -^> ExportCalendarMeetings.bas
    echo   (Russian UI uses ChrW/H - encoding-safe)
    echo.
  )
) else (
  echo Missing import-macro.vbs. Manual import:
  echo   Alt+F11 -^> delete old module -^> Import File -^> ExportCalendarMeetings.bas
)

echo.
echo ========================================
echo ALWAYS WORKS:
echo   Alt+F8 -^> ExportManagerCalendarMeetings -^> Run
echo.
echo Button: Calendar tab
echo NEVER run InstallCalendarExportButton
echo ========================================
pause
