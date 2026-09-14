Attribute VB_Name = "ExportCalendarMeetings"
Option Explicit

' Full module: install.bat assembles ExportCalendarMeetings.bas.p0-p7 (or expands gz.b64).
' Encoding-safe Russian UI uses ChrW via Private Function H().

Public Sub ExportManagerCalendarMeetings()
    MsgBox "Run install.bat first (assembles ChrW-safe module from p0-p7), then Alt+F8 again.", vbExclamation
End Sub

Private Function H(ByVal hexUtf16 As String) As String
    Dim i As Long
    Dim s As String
    s = ""
    For i = 1 To Len(hexUtf16) Step 4
        s = s & ChrW(CLng("&H" & Mid$(hexUtf16, i, 4)))
    Next i
    H = s
End Function
