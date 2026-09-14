Attribute VB_Name = "ExportCalendarMeetings"
Option Explicit

' Full module: prefer install.bat (assembles p0-p7 or expands gz.b64).
' Russian UI via ChrW/H() — encoding-safe. Re-import: delete old module first.

' STUB FALLBACK if opened without install — run install.bat.
Public Sub ExportManagerCalendarMeetings()
    MsgBox "Run install.bat first (assembles ChrW .bas from p0-p7 or gz.b64), then Alt+F8 again.", vbExclamation
End Sub
