# Выгрузка встреч календаря (Outlook 2016)

## Установка

1. Code → Download ZIP → распаковать
2. Запустить `install.bat` (закроет Outlook, поставит кнопку, **удалит** старый модуль `ExportCalendarMeetings` и импортирует новый `.bas`)
3. Открыть Outlook → Календарь

### Обновление / кракозябры вместо русского

В диалогах должен быть нормальный русский (`Диапазон дат`, `Начало периода`, `Исключения из поиска`). Если нет — скачайте ZIP снова и переустановите:

- снова `install.bat`, **или**
- `Alt+F11` → удалить `ExportCalendarMeetings` → File → Import File → `ExportCalendarMeetings.bas`

Русский UI собран через `ChrW`, поэтому UTF-8 на GitHub не ломает текст в Outlook.

## Запуск

- **Alt+F8** → `ExportManagerCalendarMeetings` → Выполнить
- или кнопка на вкладке **Календарь**

Excel → папка **Изображения**.

Не запускайте `InstallCalendarExportButton` (устаревший установщик).
