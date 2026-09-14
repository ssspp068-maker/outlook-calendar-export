Attribute VB_Name = "ExportCalendarMeetings"
Option Explicit

' Placeholder only. install.bat overwrites this file by joining
' ExportCalendarMeetings.bas.p0-p7 (full ChrW/H() Russian UI module),
' then imports the assembled .bas into Outlook VBA.
' Do NOT run the old CommandBars installer macro.

Public Sub ExportManagerCalendarMeetings()
    MsgBox "Run install.bat first (assembles module from p0-p7), then Alt+F8 again.", vbExclamation
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
