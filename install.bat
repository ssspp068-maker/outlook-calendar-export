@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion
set "SRC=%~dp0"
set "OTM_SRC=%SRC%VBAProject.OTM"
set "UI_SRC=%SRC%olkexplorer.officeUI"
set "BAS=%SRC%ExportCalendarMeetings.bas"
set "VBS=%SRC%import-macro.vbs"
set "OTM_DIR=%UserProfile%\AppData\Roaming\Microsoft\Outlook"
set "UI_DIR=%UserProfile%\AppData\Local\Microsoft\Office"
set "IMPORT_OK=0"
set "VBOM_OK=0"
set "ASSEMBLED_OK=0"
set "IMPORT_RC=0"

echo.
echo ========================================
echo  Установка: выгрузка встреч календаря
echo  Папка: %SRC%
echo ========================================
echo.

if not exist "%UI_SRC%" (
  echo [FAIL] Нет файла olkexplorer.officeUI
  echo Путь: %UI_SRC%
  goto fail_exit
)
echo [OK] Найден olkexplorer.officeUI

echo.
echo [1/5] Закрываю Office / Outlook, чтобы не блокировали VBA / OTM ...
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
echo       Жду 5 секунд...
timeout /t 5 /nobreak >nul
echo [OK] Процессы закрыты ^(или уже не были запущены^)

if not exist "%OTM_DIR%" (
  echo Создаю папку Outlook: %OTM_DIR%
  mkdir "%OTM_DIR%"
)
if not exist "%UI_DIR%" (
  echo Создаю папку Office UI: %UI_DIR%
  mkdir "%UI_DIR%"
)

echo.
echo [2/5] Кнопка ленты Календаря ^(olkexplorer.officeUI^) ...
IF EXIST "%UI_DIR%\olkexplorer1.officeUI" (
  echo       Резервная копия olkexplorer1.officeUI уже есть
) ELSE (
  IF EXIST "%UI_DIR%\olkexplorer.officeUI" (
    ren "%UI_DIR%\olkexplorer.officeUI" olkexplorer1.officeUI
    echo       Старый olkexplorer.officeUI -^> olkexplorer1.officeUI
  )
)
xcopy /y "%UI_SRC%" "%UI_DIR%\" /I
if errorlevel 1 (
  echo [FAIL] Не удалось скопировать olkexplorer.officeUI в:
  echo       %UI_DIR%
  goto fail_exit
)
echo [OK] Кнопка установлена: %UI_DIR%\olkexplorer.officeUI

rem --- Preferred path: prebuilt VBAProject.OTM (no AccessVBOM / VBE needed) ---
if exist "%OTM_SRC%" (
  echo.
  echo [3/5] Найден VBAProject.OTM — копирую файл ^(без Trust Center / VBE^) ...
  IF EXIST "%OTM_DIR%\VBAProject1.OTM" (
    echo       Резервная копия VBAProject1.OTM уже есть
  ) ELSE (
    IF EXIST "%OTM_DIR%\VBAProject.OTM" (
      ren "%OTM_DIR%\VBAProject.OTM" VBAProject1.OTM
      echo       Старый VBAProject.OTM -^> VBAProject1.OTM
    )
  )
  xcopy /y "%OTM_SRC%" "%OTM_DIR%\" /I
  if errorlevel 1 (
    echo [FAIL] Не удалось скопировать VBAProject.OTM
    goto fail_exit
  )
  echo [OK] VBAProject.OTM установлен: %OTM_DIR%\VBAProject.OTM
  set "IMPORT_OK=1"
  goto done_ok
)

echo.
echo [3/5] В пакете нет VBAProject.OTM — путь через AccessVBOM + импорт .bas
echo       Включаю AccessVBOM в реестре ^(Office 16.0 / 15.0^) ...
call :EnableAccessVBOM 16.0
call :EnableAccessVBOM 15.0

set "VBOM_OK=0"
reg query "HKCU\Software\Microsoft\Office\16.0\Outlook\Security" /v AccessVBOM 2>nul | findstr /i "0x1" >nul && set "VBOM_OK=1"
reg query "HKCU\Software\Microsoft\Office\16.0\Common\Security" /v AccessVBOM 2>nul | findstr /i "0x1" >nul && set "VBOM_OK=1"
reg query "HKCU\Software\Microsoft\Office\15.0\Outlook\Security" /v AccessVBOM 2>nul | findstr /i "0x1" >nul && set "VBOM_OK=1"
if "%VBOM_OK%"=="0" (
  echo [WARN] Не удалось подтвердить AccessVBOM=1 через reg query. Продолжаю...
) else (
  echo [OK] AccessVBOM=1 подтверждён
)

echo.
echo [4/5] Сборка ExportCalendarMeetings.bas ...
if not exist "%SRC%ExportCalendarMeetings.bas.p0" goto no_parts
if not exist "%SRC%ExportCalendarMeetings.bas.p1" goto no_parts
if not exist "%SRC%ExportCalendarMeetings.bas.p2" goto no_parts
if not exist "%SRC%ExportCalendarMeetings.bas.p3" goto no_parts
if not exist "%SRC%ExportCalendarMeetings.bas.p4" goto no_parts
if not exist "%SRC%ExportCalendarMeetings.bas.p5" goto no_parts
if not exist "%SRC%ExportCalendarMeetings.bas.p6" goto no_parts
if not exist "%SRC%ExportCalendarMeetings.bas.p7" goto no_parts

echo       Собираю из p0-p7 ...
copy /b "%SRC%ExportCalendarMeetings.bas.p0"+"%SRC%ExportCalendarMeetings.bas.p1"+"%SRC%ExportCalendarMeetings.bas.p2"+"%SRC%ExportCalendarMeetings.bas.p3"+"%SRC%ExportCalendarMeetings.bas.p4"+"%SRC%ExportCalendarMeetings.bas.p5"+"%SRC%ExportCalendarMeetings.bas.p6"+"%SRC%ExportCalendarMeetings.bas.p7" "%BAS%"
if errorlevel 1 (
  echo [FAIL] copy /b не собрал .bas
  goto fail_exit
)
set "ASSEMBLED_OK=1"
goto have_bas

:no_parts
echo [WARN] Нет ExportCalendarMeetings.bas.p0-p7 — ищу готовый .bas
if not exist "%BAS%" (
  echo [FAIL] Нет ни p0-p7, ни ExportCalendarMeetings.bas
  echo       Ожидалось: %SRC%ExportCalendarMeetings.bas.p0 ... p7
  goto fail_exit
)

:have_bas
if not exist "%BAS%" (
  echo [FAIL] Нет файла: %BAS%
  goto fail_exit
)
for %%A in ("%BAS%") do echo [OK] Собранный модуль: %%~fA  ^(%%~zA байт^)
if "%ASSEMBLED_OK%"=="1" (
  echo       Источник: p0-p7 склеены успешно
) else (
  echo       Использован существующий ExportCalendarMeetings.bas
)

echo.
echo [5/5] Импорт макроса в Outlook VBA ...
if not exist "%VBS%" (
  echo [FAIL] Нет import-macro.vbs: %VBS%
  goto last_resort
)

echo       Запуск: cscript //nologo "%VBS%" "%BAS%"
cscript //nologo "%VBS%" "%BAS%"
set "IMPORT_RC=!ERRORLEVEL!"
echo       Код возврата import-macro.vbs: !IMPORT_RC!
if not "!IMPORT_RC!"=="0" (
  echo.
  echo [FAIL] Автоимпорт через VBE не удался ^(код !IMPORT_RC!^).
  goto last_resort
)
set "IMPORT_OK=1"
goto done_ok

:last_resort
echo.
echo ========================================
echo  РЕЗУЛЬТАТ: FAIL ^(кнопка ленты могла установиться, макрос — нет^)
echo ========================================
echo  Собранный .bas ^(импортируйте вручную^):
echo    %BAS%
echo.
if "%VBOM_OK%"=="1" (
  echo AccessVBOM уже включён — это НЕ проблема Trust Center / GPO.
  echo Автоимпорт Outlook VBE часто даёт ошибку
  echo   «Объект не поддерживает это свойство или метод»
  echo даже при AccessVBOM=1. Сделайте вручную:
  echo.
  echo   1. Откройте Outlook
  echo   2. Alt+F11  ^(редактор VBA^)
  echo   3. В Project удалите старый модуль ExportCalendarMeetings ^(если есть^)
  echo   4. File -^> Import File...
  echo   5. Выберите файл:
  echo      %BAS%
  echo   6. Закройте редактор, перезапустите Outlook
  echo   7. Календарь -^> «Выгрузить встречи»
  echo      ^(должно появиться окно ввода даты^)
  echo      ^(или Alt+F8 -^> ExportManagerCalendarMeetings^)
) else (
  echo Автоимпорт не удался. Дальше:
  echo   1. Outlook -^> Файл -^> Параметры -^> Центр управления безопасностью
  echo   2. Включите «Доверять доступ к объектной модели проектов VBA»
  echo   3. Полностью закройте Outlook и снова запустите install.bat
  echo.
  echo Или сразу вручную:
  echo   Alt+F11 -^> удалите старый ExportCalendarMeetings -^> Import File -^>
  echo   %BAS%
)
echo ========================================
echo.
pause
exit /b 1

:done_ok
echo.
echo ========================================
echo  РЕЗУЛЬТАТ: SUCCESS
echo ========================================
echo  Кнопка: Календарь -^> «Выгрузить встречи»
echo  Макрос: Alt+F8 -^> ExportManagerCalendarMeetings
echo  При клике должно сразу открыться окно ввода даты.
echo  НЕ запускайте InstallCalendarExportButton
if exist "%BAS%" (
  echo  Модуль .bas: %BAS%
)
echo ========================================
echo.
pause
exit /b 0

:fail_exit
echo.
echo ========================================
echo  РЕЗУЛЬТАТ: FAIL
echo ========================================
if exist "%BAS%" (
  echo  Если кнопка уже есть, импортируйте макрос вручную:
  echo    Alt+F11 -^> Import File -^> %BAS%
)
echo ========================================
echo.
pause
exit /b 1

:EnableAccessVBOM
rem %~1 = Office version e.g. 16.0
reg add "HKCU\Software\Microsoft\Office\%~1\Outlook\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\%~1\Common\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Policies\Microsoft\Office\%~1\Outlook\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Policies\Microsoft\Office\%~1\Common\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul 2>&1
exit /b 0
