# Выгрузка встреч календаря (Outlook 2016)

## Если debug / break mode на старом InstallCalendarExportButton

1. VBA → **Reset** (■ / Run → Reset)
2. `Alt+F11` → удалите модуль `ExportCalendarMeetings` (Remove → No)
3. Закройте Outlook и запустите `install.bat` (развернёт `.bas` из `.gz.b64` и импортирует)
4. `Alt+F8` → `ExportManagerCalendarMeetings` → Выполнить
5. **НИКОГДА** не запускайте `InstallCalendarExportButton`

## Установка

1. Скачайте ZIP с GitHub (`Code` → Download ZIP) и распакуйте
2. Запустите `install.bat`
3. Outlook → `Alt+F8` → `ExportManagerCalendarMeetings`

Перед повторным импортом удалите старый модуль `ExportCalendarMeetings` (или дайте `install.bat` сделать это сам).

Русский UI — через `ChrW`/`H()` (не зависит от кодировки `.bas`).
