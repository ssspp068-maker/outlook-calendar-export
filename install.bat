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

REM Always rebuild from part0+1+2 when present (overrides corrupt/stale .gz.b64).
if exist "%SRC%ExportCalendarMeetings.bas.gz.b64.part0" if exist "%SRC%ExportCalendarMeetings.bas.gz.b64.part1" if exist "%SRC%ExportCalendarMeetings.bas.gz.b64.part2" (
  echo Joining bas.gz.b64 parts ...
  copy /b "%SRC%ExportCalendarMeetings.bas.gz.b64.part0"+"%SRC%ExportCalendarMeetings.bas.gz.b64.part1"+"%SRC%ExportCalendarMeetings.bas.gz.b64.part2" "%B64%" >nul
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

if not exist "%BAS%" (
  echo [ERROR] Missing ExportCalendarMeetings.bas
  pause
  exit /b 1
)

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

echo Importing macro (removes old ExportCalendarMeetings, then Import)...
if exist "%VBS%" (
  cscript //nologo "%VBS%" "%BAS%"
  if errorlevel 1 (
    echo.
    echo Auto-import failed. Manual once:
    echo   Outlook -^> Alt+F11 -^> delete ExportCalendarMeetings if present
    echo   File -^> Import File -^> ExportCalendarMeetings.bas
    echo   (Cyrillic UI uses ChrW/H — encoding-safe)
    echo.
  )
) else (
  echo Missing import-macro.vbs. Manual import:
  echo   Outlook -^> Alt+F11 -^> delete ExportCalendarMeetings if present
  echo   File -^> Import File -^> ExportCalendarMeetings.bas
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
