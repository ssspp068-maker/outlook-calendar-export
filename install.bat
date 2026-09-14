@echo off
chcp 65001 >nul
setlocal
set "SRC=%~dp0"
set "UI_LOCAL=%UserProfile%\AppData\Local\Microsoft\Office"
set "UI_ROAM=%AppData%\Microsoft\Office"
set "BAS=%SRC%ExportCalendarMeetings.bas"
set "VBS=%SRC%import-macro.vbs"

if not exist "%BAS%" (
  echo [ОШИБКА] Нет ExportCalendarMeetings.bas рядом с install.bat
  pause
  exit /b 1
)
if not exist "%SRC%olkexplorer.officeUI" (
  echo [ОШИБКА] Нет olkexplorer.officeUI рядом с install.bat
  pause
  exit /b 1
)

echo.
echo ========================================
echo ЕСЛИ БЫЛ DEBUG / СТАРЫЙ МОДУЛЬ:
echo   1^) Outlook VBA → Reset (■^), закройте Outlook
echo   2^) install.bat удалит старый ExportCalendarMeetings сам
echo   3^) НЕ запускайте InstallCalendarExportButton
echo      (CommandBars-установщик удалён^)
echo ========================================
echo.

echo Закрываю Outlook...
taskkill /f /im OUTLOOK.EXE >nul 2>&1
timeout /t 2 /nobreak >nul

reg add "HKCU\Software\Microsoft\Office\16.0\Outlook\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\Common\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul

if not exist "%UI_LOCAL%" mkdir "%UI_LOCAL%"
if not exist "%UI_ROAM%" mkdir "%UI_ROAM%"

if not exist "%UI_LOCAL%\olkexplorer1.officeUI" if exist "%UI_LOCAL%\olkexplorer.officeUI" (
  ren "%UI_LOCAL%\olkexplorer.officeUI" olkexplorer1.officeUI
)
if not exist "%UI_ROAM%\olkexplorer1.officeUI" if exist "%UI_ROAM%\olkexplorer.officeUI" (
  ren "%UI_ROAM%\olkexplorer.officeUI" olkexplorer1.officeUI
)

xcopy /y "%SRC%olkexplorer.officeUI" "%UI_LOCAL%\" /I >nul
xcopy /y "%SRC%olkexplorer.officeUI" "%UI_ROAM%\" /I >nul
echo Кнопка (файл ленты)... OK

echo Импортирую макрос (старый ExportCalendarMeetings удаляется, затем Import)...
cscript //nologo "%VBS%" "%BAS%"
set "ERR=%ERRORLEVEL%"

if not "%ERR%"=="0" (
  echo.
  echo Автоимпорт не вышел. Вручную:
  echo   Outlook → Alt+F11 → удалить модуль ExportCalendarMeetings ^(если есть^)
  echo   File → Import File → ExportCalendarMeetings.bas
  echo   ^(кириллица в UI: строки через ChrW^)
  echo.
)

echo.
echo ========================================
echo ЗАПУСК (единственный макрос):
echo   Alt+F8 → ExportManagerCalendarMeetings → Выполнить
echo.
echo НЕ запускайте InstallCalendarExportButton — его нет в модуле.
echo.
echo Кнопка на «Календарь» вручную (по желанию):
echo   1^) Откройте Календарь
echo   2^) Файл → Параметры → Настроить ленту
echo   3^) Справа: Календарь → Создать группу
echo   4^) Слева: Макросы → ExportManagerCalendarMeetings → Добавить
echo   5^) ОК
echo ========================================
pause
