1. Скачайте ZIP (Code → Download ZIP) и распакуйте.
2. Запустите `install.bat` (окно покажет шаги SUCCESS/FAIL и путь к `.bas`; закроет Outlook/Office).
3. Откройте Outlook → **Календарь** → «Выгрузить встречи» — должно сразу появиться окно ввода даты (или Alt+F8 → `ExportManagerCalendarMeetings`).

Если автоимпорт не удался: Alt+F11 → удалите старый модуль `ExportCalendarMeetings` → Import File → путь к `.bas` из окна `install.bat`.

Не запускайте `InstallCalendarExportButton`.
