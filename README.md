# Выгрузка встреч календаря (Outlook 2016)

## Установка

1. Code → Download ZIP → распаковать
2. Запустить `install.bat` (соберёт `.bas` из `*.gz.b64.part*`, импортирует макрос, **удалит** старый модуль)
3. Открыть Outlook → Календарь

### Кракозябры вместо русского

Должно быть: «Диапазон дат», «Начало периода», «Исключения из поиска». Если нет:

- снова `install.bat`, или
- `Alt+F11` → удалить `ExportCalendarMeetings` → Import `ExportCalendarMeetings.bas`

Русский UI через `ChrW`/`H()` — не зависит от кодировки файла на GitHub.

## Запуск

`Alt+F8` → `ExportManagerCalendarMeetings`

Excel → **Изображения**. Не запускайте `InstallCalendarExportButton`.
