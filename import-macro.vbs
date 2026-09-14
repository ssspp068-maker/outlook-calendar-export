Option Explicit

Dim basPath, ol, ns, vbProj, vbComp, i, errNum, errDesc

If WScript.Arguments.Count < 1 Then
  WScript.Echo "Need path to .bas file"
  WScript.Quit 1
End If

basPath = WScript.Arguments(0)

On Error Resume Next

Set ol = CreateObject("Outlook.Application")
errNum = Err.Number: errDesc = Err.Description: Err.Clear
If errNum <> 0 Or ol Is Nothing Then
  WScript.Echo "Cannot start Outlook: " & errDesc
  WScript.Quit 1
End If

Set ns = ol.GetNamespace("MAPI")
ns.Logon "", "", False, False
Err.Clear

WScript.Sleep 2000

Set vbProj = ol.VBE.ActiveVBProject
errNum = Err.Number: errDesc = Err.Description: Err.Clear
If errNum <> 0 Or vbProj Is Nothing Then
  WScript.Echo "No access to Outlook VBA project."
  WScript.Echo "Enable: File > Options > Trust Center > Trust Center Settings >"
  WScript.Echo "Macro Settings > Trust access to the VBA project object model"
  On Error Resume Next
  ol.Quit
  WScript.Quit 2
End If

For i = vbProj.VBComponents.Count To 1 Step -1
  Set vbComp = vbProj.VBComponents.Item(i)
  If Not vbComp Is Nothing Then
    If StrComp(vbComp.Name, "ExportCalendarMeetings", vbTextCompare) = 0 Then
      vbProj.VBComponents.Remove vbComp
      Err.Clear
    ElseIf StrComp(vbComp.Name, "RuUiStrings", vbTextCompare) = 0 Then
      vbProj.VBComponents.Remove vbComp
      Err.Clear
    End If
  End If
Next

vbProj.VBComponents.Import basPath
errNum = Err.Number: errDesc = Err.Description: Err.Clear
If errNum <> 0 Then
  WScript.Echo "Import failed: " & errDesc
  ol.Quit
  WScript.Quit 3
End If

WScript.Echo "Module ExportCalendarMeetings imported."

ol.Quit
WScript.Sleep 1500
WScript.Quit 0
