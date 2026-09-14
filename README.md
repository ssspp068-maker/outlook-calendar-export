# Выгрузка встреч календаря (Outlook 2016)

## Если debug / break mode на старом InstallCalendarExportButton

1. VBA → **Reset** (■ / Run → Reset)
2. `Alt+F11` → удалите модуль `ExportCalendarMeetings` (Remove → No)
3. Import нового `ExportCalendarMeetings.bas` или закройте Outlook и запустите `install.bat`
4. `Alt+F8` → `ExportManagerCalendarMeetings` → Выполнить
5. **НИКОГДА** не запускайте `InstallCalendarExportButton`

## Установка

1. Скачайте ZIP (Code → Download ZIP) и распакуйте.
2. Запустите `install.bat` (закроет Outlook, скопирует кнопку, импортирует макрос).
3. Откройте **Календарь** → кнопка «Выгрузить встречи», или всегда: **Alt+F8** → `ExportManagerCalendarMeetings`.
4. Если старый модуль мешает: **Alt+F11** → удалите `ExportCalendarMeetings` → снова `install.bat` или Import `.bas`.

Русский UI — через `ChrW`/`CW` (не зависит от кодировки `.bas`).
