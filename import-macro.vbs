Option Explicit

' Import ExportCalendarMeetings.bas into Outlook's VBA project via VBE.
' Prefer VBProjects(1) over ActiveVBProject (often 438 under automation).

Dim basPath, ol, ns, vbe, vbProj, vbComp, i, errNum, errDesc, sh
Dim stepName, explorer, folder, usedGetObject

If WScript.Arguments.Count < 1 Then
  WScript.Echo "Need path to .bas file"
  WScript.Quit 1
End If

basPath = WScript.Arguments(0)
usedGetObject = False
stepName = "init"

On Error Resume Next

' Force AccessVBOM before any Outlook COM start (install.bat should already set this).
stepName = "RegWrite AccessVBOM"
Set sh = CreateObject("WScript.Shell")
sh.RegWrite "HKCU\Software\Microsoft\Office\16.0\Outlook\Security\AccessVBOM", 1, "REG_DWORD"
sh.RegWrite "HKCU\Software\Microsoft\Office\16.0\Common\Security\AccessVBOM", 1, "REG_DWORD"
sh.RegWrite "HKCU\Software\Microsoft\Office\15.0\Outlook\Security\AccessVBOM", 1, "REG_DWORD"
sh.RegWrite "HKCU\Software\Microsoft\Office\15.0\Common\Security\AccessVBOM", 1, "REG_DWORD"
Err.Clear

' Prefer an already-running Outlook so its VBA project is loaded.
stepName = "GetObject Outlook.Application"
Set ol = GetObject(, "Outlook.Application")
errNum = Err.Number: errDesc = Err.Description: Err.Clear
If errNum <> 0 Or ol Is Nothing Then
  stepName = "CreateObject Outlook.Application"
  Set ol = CreateObject("Outlook.Application")
  errNum = Err.Number: errDesc = Err.Description: Err.Clear
  If errNum <> 0 Or ol Is Nothing Then
    WScript.Echo "FAIL at " & stepName & ": [" & errNum & "] " & errDesc
    WScript.Quit 1
  End If
Else
  usedGetObject = True
End If

stepName = "GetNamespace MAPI"
Set ns = ol.GetNamespace("MAPI")
errNum = Err.Number: errDesc = Err.Description: Err.Clear
If errNum <> 0 Or ns Is Nothing Then
  WScript.Echo "FAIL at " & stepName & ": [" & errNum & "] " & errDesc
  Call SoftQuit
  WScript.Quit 1
End If

stepName = "MAPI Logon"
ns.Logon "", "", False, False
Err.Clear

' Make an Explorer visible so Outlook loads VBAProject.OTM / VBE.
stepName = "Ensure Explorer"
If ol.Explorers.Count = 0 Then
  ' 6 = olFolderInbox — reliable default to open an Explorer window
  Set folder = ns.GetDefaultFolder(6)
  errNum = Err.Number: errDesc = Err.Description: Err.Clear
  If errNum <> 0 Or folder Is Nothing Then
    WScript.Echo "FAIL at GetDefaultFolder: [" & errNum & "] " & errDesc
    Call SoftQuit
    WScript.Quit 1
  End If
  Set explorer = folder.GetExplorer()
  If Not explorer Is Nothing Then
    explorer.Display
  End If
  errNum = Err.Number: errDesc = Err.Description: Err.Clear
  If errNum <> 0 Then
    WScript.Echo "FAIL at Explorer.Display: [" & errNum & "] " & errDesc
    Call SoftQuit
    WScript.Quit 1
  End If
End If
WScript.Sleep 3000

stepName = "ol.VBE"
Set vbe = ol.VBE
errNum = Err.Number: errDesc = Err.Description: Err.Clear
If errNum <> 0 Or vbe Is Nothing Then
  WScript.Echo "FAIL at " & stepName & ": [" & errNum & "] " & errDesc
  WScript.Echo "AccessVBOM may be off, or VBE is unavailable for this Outlook session."
  Call SoftQuit
  WScript.Quit 2
End If

' Prefer VBProjects(1) / iterate — ActiveVBProject often throws 438 under automation.
stepName = "VBProjects"
Set vbProj = Nothing
Err.Clear
If vbe.VBProjects.Count >= 1 Then
  Set vbProj = vbe.VBProjects.Item(1)
  errNum = Err.Number: errDesc = Err.Description: Err.Clear
  If errNum <> 0 Then
    WScript.Echo "WARN VBProjects(1): [" & errNum & "] " & errDesc
    Set vbProj = Nothing
    Err.Clear
  End If
End If

If vbProj Is Nothing Then
  For i = 1 To vbe.VBProjects.Count
    Set vbProj = vbe.VBProjects.Item(i)
    errNum = Err.Number: Err.Clear
    If errNum = 0 And Not vbProj Is Nothing Then Exit For
    Set vbProj = Nothing
  Next
End If

If vbProj Is Nothing Then
  stepName = "ActiveVBProject (fallback)"
  Set vbProj = vbe.ActiveVBProject
  errNum = Err.Number: errDesc = Err.Description: Err.Clear
  If errNum <> 0 Or vbProj Is Nothing Then
    WScript.Echo "FAIL at " & stepName & ": [" & errNum & "] " & errDesc
    WScript.Echo "Could not resolve any Outlook VBProject (not a domain GPO issue if AccessVBOM=1)."
    Call SoftQuit
    WScript.Quit 2
  End If
End If

WScript.Echo "Using VBProject: " & vbProj.Name

stepName = "Remove old modules"
For i = vbProj.VBComponents.Count To 1 Step -1
  Set vbComp = vbProj.VBComponents.Item(i)
  If Not vbComp Is Nothing Then
    If StrComp(vbComp.Name, "ExportCalendarMeetings", vbTextCompare) = 0 Or _
       StrComp(vbComp.Name, "RuUiStrings", vbTextCompare) = 0 Then
      vbProj.VBComponents.Remove vbComp
      Err.Clear
    End If
  End If
Next
Err.Clear

stepName = "Import"
vbProj.VBComponents.Import basPath
errNum = Err.Number: errDesc = Err.Description: Err.Clear
If errNum <> 0 Then
  WScript.Echo "FAIL at " & stepName & ": [" & errNum & "] " & errDesc
  WScript.Echo "Manual fix: Alt+F11 -> File -> Import File -> ExportCalendarMeetings.bas"
  Call SoftQuit
  WScript.Quit 3
End If

WScript.Echo "Module ExportCalendarMeetings imported."
If usedGetObject Then
  WScript.Echo "Attached to existing Outlook; quitting to flush VBAProject.OTM ..."
Else
  WScript.Echo "Quitting Outlook to flush VBAProject.OTM ..."
End If

' Quit so Outlook writes VBAProject.OTM to disk.
WScript.Sleep 1500
stepName = "Quit"
ol.Quit
Err.Clear
WScript.Sleep 2500
WScript.Quit 0

Sub SoftQuit
  On Error Resume Next
  If Not ol Is Nothing Then ol.Quit
  WScript.Sleep 1000
End Sub
