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
set "VBOM_OK=0"

if not exist "%UI_SRC%" (
  echo [ERROR] Нет файла olkexplorer.officeUI
  pause
  exit /b 1
)

echo Закрываю Office / Outlook, чтобы не блокировали VBA / OTM ...
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
timeout /t 5 /nobreak >nul

if not exist "%OTM_DIR%" mkdir "%OTM_DIR%"
if not exist "%UI_DIR%" mkdir "%UI_DIR%"

IF EXIST "%UI_DIR%\olkexplorer1.officeUI" (
  ECHO Резервная копия olkexplorer1.officeUI уже есть
) ELSE (
  IF EXIST "%UI_DIR%\olkexplorer.officeUI" ren "%UI_DIR%\olkexplorer.officeUI" olkexplorer1.officeUI
)

echo Копирую кнопку ленты Календаря ...
xcopy /y "%UI_SRC%" "%UI_DIR%\" /I >nul

rem --- Preferred path: prebuilt VBAProject.OTM (no AccessVBOM / VBE needed) ---
if exist "%OTM_SRC%" (
  echo Найден VBAProject.OTM — ставлю копированием файла ^(без Trust Center / VBE^) ...
  IF EXIST "%OTM_DIR%\VBAProject1.OTM" (
    ECHO Резервная копия VBAProject1.OTM уже есть
  ) ELSE (
    IF EXIST "%OTM_DIR%\VBAProject.OTM" ren "%OTM_DIR%\VBAProject.OTM" VBAProject1.OTM
  )
  xcopy /y "%OTM_SRC%" "%OTM_DIR%\" /I >nul
  if errorlevel 1 (
    echo [ERROR] Не удалось скопировать VBAProject.OTM
    pause
    exit /b 1
  )
  echo VBAProject.OTM установлен.
  goto done_ok
)

rem --- Fallback: force AccessVBOM in registry, then import .bas via VBE ---
echo В пакете нет VBAProject.OTM — включаю AccessVBOM в реестре и импортирую .bas ...
call :EnableAccessVBOM 16.0
call :EnableAccessVBOM 15.0

echo Проверяю AccessVBOM ...
set "VBOM_OK=0"
reg query "HKCU\Software\Microsoft\Office\16.0\Outlook\Security" /v AccessVBOM 2>nul | findstr /i "0x1" >nul && set "VBOM_OK=1"
reg query "HKCU\Software\Microsoft\Office\16.0\Common\Security" /v AccessVBOM 2>nul | findstr /i "0x1" >nul && set "VBOM_OK=1"
reg query "HKCU\Software\Microsoft\Office\15.0\Outlook\Security" /v AccessVBOM 2>nul | findstr /i "0x1" >nul && set "VBOM_OK=1"
if "%VBOM_OK%"=="0" (
  echo [WARN] Не удалось подтвердить AccessVBOM=1 через reg query. Продолжаю...
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

echo Собираю ExportCalendarMeetings.bas из p0-p7 ...
copy /b "%SRC%ExportCalendarMeetings.bas.p0"+"%SRC%ExportCalendarMeetings.bas.p1"+"%SRC%ExportCalendarMeetings.bas.p2"+"%SRC%ExportCalendarMeetings.bas.p3"+"%SRC%ExportCalendarMeetings.bas.p4"+"%SRC%ExportCalendarMeetings.bas.p5"+"%SRC%ExportCalendarMeetings.bas.p6"+"%SRC%ExportCalendarMeetings.bas.p7" "%BAS%" >nul
goto have_bas

:no_parts
if not exist "%BAS%" (
  echo [ERROR] Нет ExportCalendarMeetings.bas.p0-p7
  pause
  exit /b 1
)

:have_bas
if not exist "%BAS%" (
  echo [ERROR] Нет ExportCalendarMeetings.bas
  pause
  exit /b 1
)

if not exist "%VBS%" (
  echo [ERROR] Нет import-macro.vbs
  goto last_resort
)

echo Импортирую макрос в Outlook VBA ...
cscript //nologo "%VBS%" "%BAS%"
if errorlevel 1 (
  echo.
  echo Автоимпорт через VBE не удался.
  goto last_resort
)
set "IMPORT_OK=1"
goto done_ok

:last_resort
echo.
echo ========================================
if "%VBOM_OK%"=="1" (
  echo AccessVBOM уже включён — это НЕ проблема Trust Center / GPO.
  echo Автоимпорт Outlook VBE часто даёт ошибку
  echo   «Объект не поддерживает это свойство или метод»
  echo даже при AccessVBOM=1. Сделайте вручную:
  echo.
  echo   1. Откройте Outlook
  echo   2. Alt+F11  ^(редактор VBA^)
  echo   3. File -^> Import File...
  echo   4. Выберите файл:
  echo      %BAS%
  echo   5. Закройте редактор, перезапустите Outlook
  echo   6. Календарь -^> «Выгрузить встречи»
  echo      ^(или Alt+F8 -^> ExportManagerCalendarMeetings^)
) else (
  echo Автоимпорт не удался. Дальше:
  echo   1. Outlook -^> Файл -^> Параметры -^> Центр управления безопасностью
  echo   2. Включите «Доверять доступ к объектной модели проектов VBA»
  echo   3. Полностью закройте Outlook и снова запустите install.bat
  echo.
  echo Или сразу вручную: Alt+F11 -^> Import File -^> ExportCalendarMeetings.bas
)
echo ========================================
echo.
pause
exit /b 1

:done_ok
echo.
echo ========================================
echo OK. Откройте Outlook -^> Календарь -^> «Выгрузить встречи»
echo Или всегда: Alt+F8 -^> ExportManagerCalendarMeetings
echo НЕ запускайте InstallCalendarExportButton
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
