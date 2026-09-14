@echo off
chcp 65001 >nul
setlocal
set "SRC=%~dp0"
set "UI=%UserProfile%\AppData\Local\Microsoft\Office"
set "BAS=%SRC%ExportCalendarMeetings.bas"
set "VBS=%SRC%import-macro.vbs"

if not exist "%BAS%" (
  echo [ОШИБКА] Нет файла ExportCalendarMeetings.bas рядом с install.bat
  pause
  exit /b 1
)
if not exist "%SRC%olkexplorer.officeUI" (
  echo [ОШИБКА] Нет файла olkexplorer.officeUI рядом с install.bat
  pause
  exit /b 1
)

echo Закрываю Outlook...
taskkill /f /im OUTLOOK.EXE >nul 2>&1
timeout /t 2 /nobreak >nul

REM Доступ к VBA-проекту (Outlook 2016 = 16.0) — нужно для автоимпорта модуля
reg add "HKCU\Software\Microsoft\Office\16.0\Outlook\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\Common\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul

REM Кнопка на вкладке «Календарь»
if not exist "%UI%" mkdir "%UI%"
if not exist "%UI%\olkexplorer1.officeUI" if exist "%UI%\olkexplorer.officeUI" (
  ren "%UI%\olkexplorer.officeUI" olkexplorer1.officeUI
)
xcopy /y "%SRC%olkexplorer.officeUI" "%UI%\" /I >nul

echo Ставлю кнопку... OK
echo Импортирую макрос...
cscript //nologo "%VBS%" "%BAS%"
set "ERR=%ERRORLEVEL%"

if not "%ERR%"=="0" (
  echo.
  echo Автоимпорт не вышел. Сделайте один раз вручную:
  echo   Outlook → Alt+F11 → Import File → ExportCalendarMeetings.bas
  echo.
  echo Если пишет про доступ к VBA: включите галку
  echo   «Доверять доступ к объектной модели проектов VBA»
  echo   в параметрах макросов Outlook.
  pause
  exit /b %ERR%
)

echo.
echo Готово. Откройте Outlook → Календарь → «Выгрузить встречи»
pause
