# outlook-calendar-export

Outlook macro: export manager meetings to Excel.

## Install

1. Download ZIP (Code -> Download ZIP) and unpack.
2. Run `install.bat` (closes Outlook, copies ribbon button, imports macro).
3. Calendar tab button, or always: **Alt+F8** -> `ExportManagerCalendarMeetings`.

## Re-import (Cyrillic fix)

If InputBox/MsgBox shows mojibake instead of Russian:

1. **Alt+F11** -> delete module `ExportCalendarMeetings` (Remove -> No).
2. Run `install.bat` again, or File -> Import File -> `ExportCalendarMeetings.bas`.

Russian UI is built with `ChrW`/`H()` (ASCII-safe on GitHub). Expected prompts: «Диапазон дат», «Начало периода…», «Исключения из поиска».

**NEVER** run `InstallCalendarExportButton`.
