@echo off
chcp 65001 >nul
setlocal
set "SRC=%~dp0"
set "UI_DIR=%UserProfile%\AppData\Local\Microsoft\Office"
set "BAS=%SRC%ExportCalendarMeetings.bas"
set "VBS=%SRC%import-macro.vbs"
set "UI=%SRC%olkexplorer.officeUI"

if not exist "%BAS%" (
  echo [ERROR] Missing ExportCalendarMeetings.bas
  pause
  exit /b 1
)
if not exist "%UI%" (
  echo [ERROR] Missing olkexplorer.officeUI
  pause
  exit /b 1
)

echo Closing Outlook...
taskkill /f /im OUTLOOK.EXE >nul 2>&1
timeout /t 2 /nobreak >nul

rem Allow VBS to import into VBA project
reg add "HKCU\Software\Microsoft\Office\16.0\Outlook\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\Common\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul

if not exist "%UI_DIR%" mkdir "%UI_DIR%"

rem Backup existing ribbon UI once (same pattern as working install example)
IF EXIST "%UI_DIR%\olkexplorer1.officeUI" (
  ECHO Backup olkexplorer1.officeUI already exists
) ELSE (
  IF EXIST "%UI_DIR%\olkexplorer.officeUI" ren "%UI_DIR%\olkexplorer.officeUI" olkexplorer1.officeUI
)

xcopy /y "%UI%" "%UI_DIR%\" /I
echo Ribbon button copied to Local\Microsoft\Office

echo Importing macro into existing VBAProject.OTM ...
if exist "%VBS%" (
  cscript //nologo "%VBS%" "%BAS%"
  if errorlevel 1 (
    echo.
    echo Auto-import failed. Manual once:
    echo   Outlook -^> Alt+F11 -^> File -^> Import File -^> ExportCalendarMeetings.bas
    echo.
  )
) else (
  echo Missing import-macro.vbs. Manual import:
  echo   Outlook -^> Alt+F11 -^> File -^> Import File -^> ExportCalendarMeetings.bas
)

echo.
echo ========================================
echo ALWAYS WORKS:
echo   Alt+F8 -^> ExportManagerCalendarMeetings -^> Run
echo.
echo Button: Calendar tab, group Выгрузка
echo.
echo NEVER run the old CommandBars installer macro
echo ========================================
pause
