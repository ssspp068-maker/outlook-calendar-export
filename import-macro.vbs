Option Explicit

Dim basPath, ol, ns, vbProj, vbComp, i, errNum, errDesc

If WScript.Arguments.Count < 1 Then
  WScript.Echo "Нужен путь к .bas"
  WScript.Quit 1
End If

basPath = WScript.Arguments(0)

On Error Resume Next

Set ol = CreateObject("Outlook.Application")
errNum = Err.Number: errDesc = Err.Description: Err.Clear
If errNum <> 0 Or ol Is Nothing Then
  WScript.Echo "Не удалось запустить Outlook: " & errDesc
  WScript.Quit 1
End If

Set ns = ol.GetNamespace("MAPI")
ns.Logon "", "", False, False
Err.Clear

' Дать Outlook догрузиться
WScript.Sleep 2000

Set vbProj = ol.VBE.ActiveVBProject
errNum = Err.Number: errDesc = Err.Description: Err.Clear
If errNum <> 0 Or vbProj Is Nothing Then
  WScript.Echo "Нет доступа к VBA-проекту Outlook."
  WScript.Echo "Включите: Параметры → Центр управления безопасностью →"
  WScript.Echo "Параметры макросов → «Доверять доступ к объектной модели проектов VBA»"
  On Error Resume Next
  ol.Quit
  WScript.Quit 2
End If

' Удалить старую версию модуля, если уже была
For i = vbProj.VBComponents.Count To 1 Step -1
  Set vbComp = vbProj.VBComponents.Item(i)
  If Not vbComp Is Nothing Then
    If StrComp(vbComp.Name, "ExportCalendarMeetings", vbTextCompare) = 0 Then
      vbProj.VBComponents.Remove vbComp
      Err.Clear
    End If
  End If
Next

vbProj.VBComponents.Import basPath
errNum = Err.Number: errDesc = Err.Description: Err.Clear
If errNum <> 0 Then
  WScript.Echo "Импорт .bas не удался: " & errDesc
  ol.Quit
  WScript.Quit 3
End If

WScript.Echo "Макрос ExportCalendarMeetings импортирован."

' Закрыть Outlook — чтобы VBAProject.OTM сохранился на диск
ol.Quit
WScript.Sleep 1500
WScript.Quit 0
