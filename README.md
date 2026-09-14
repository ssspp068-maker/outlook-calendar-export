# Выгрузка встреч календаря (Outlook 2016)

## Если debug / break mode на старом InstallCalendarExportButton

1. VBA → **Reset** (■ / Run → Reset)
2. `Alt+F11` → удалите модуль `ExportCalendarMeetings` (Remove → No)
3. Закройте Outlook и запустите `install.bat`
4. `Alt+F8` → `ExportManagerCalendarMeetings` → Выполнить
5. **НИКОГДА** не запускайте `InstallCalendarExportButton`

## Установка

1. Code → Download ZIP → распаковать
2. Запустить `install.bat` (соберёт `.bas` из `part0+1+2`, **удалит** старый модуль, импортирует)
3. Outlook → `Alt+F8` → `ExportManagerCalendarMeetings`

Перед повторным импортом удалите старый модуль (или дайте `install.bat` сделать это).

Русский UI через `ChrW`/`H()` — не зависит от кодировки `.bas`.
