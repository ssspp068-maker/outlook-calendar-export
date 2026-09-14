Attribute VB_Name = "ExportCalendarMeetings"
Option Explicit

' Outlook 2016: выгрузка встреч выбранного календаря в Excel.
' Запуск: кнопка на вкладке «Календарь» (ставится install.bat + olkexplorer.officeUI)
'         или Alt+F8 → ExportManagerCalendarMeetings
'
' Фильтры:
'   - слово «ДР» заглавными в теме
'   - только встречи с участниками
'   - исключения тем: предложения + свой список (через ; или ,)
' Колонки (рус.): название, дата, место, ссылка, длительность, создатель,
'   обязательные, необязательные, вложения, число вложений
' Ссылки: Teams / Яндекс Телемост / ivavks.rosatom.ru — полный URL с путём и id
' Файл: папка «Изображения» пользователя
' Установка: install.bat (кнопка + автоимпорт .bas).

Private Const xlOpenXMLWorkbook As Long = 51

' Типовые исключения (подставляются в диалог как предложения)
Private Const DEFAULT_EXCLUSIONS As String = _
    "обед;ужин;lunch;dinner;focus time;focus;block;занят;фокус;out of office;ooo;нн"

Public Sub ExportManagerCalendarMeetings()
    Dim calendarFolder As Outlook.MAPIFolder
    Dim startDate As Date
    Dim endDate As Date
    Dim calItems As Outlook.Items
    Dim filtered As Outlook.Items
    Dim appt As Object
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim row As Long
    Dim count As Long
    Dim restrictFilter As String
    Dim savePath As String
    Dim exclusions() As String
    Dim subjectText As String
    Dim attachNames As String
    Dim attachCount As Long
    Dim meetingLink As String

    If Not IsOutlook2016Family() Then
        If MsgBox("Макрос рассчитан на Outlook 2016. Продолжить?", _
                  vbYesNo Or vbQuestion, "Outlook") = vbNo Then
            Exit Sub
        End If
    End If

    Set calendarFolder = PickCalendarFolder()
    If calendarFolder Is Nothing Then Exit Sub

    If Not AskDateRange(startDate, endDate) Then Exit Sub

    If Not AskSubjectExclusions(exclusions) Then Exit Sub

    Set calItems = calendarFolder.Items
    calItems.Sort "[Start]"
    calItems.IncludeRecurrences = True

    restrictFilter = "[Start] >= " & QuoteOutlookDate(startDate) & _
                     " AND [End] <= " & QuoteOutlookDate(DateAdd("d", 1, endDate))

    Set filtered = calItems.Restrict(restrictFilter)
    filtered.Sort "[Start]"

    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0
    If xlApp Is Nothing Then
        MsgBox "Не удалось открыть Excel. Для Outlook 2016 нужен установленный Excel.", _
               vbExclamation
        Exit Sub
    End If

    xlApp.Visible = True
    xlApp.DisplayAlerts = False
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    On Error Resume Next
    xlSheet.Name = "Встречи"
    On Error GoTo 0

    xlSheet.Range("A1:J1").Value = Array( _
        "Название", "Дата", "Место", "Ссылка", "Длительность", "Создатель", _
        "Обязательные", "Необязательные", "Вложения", "Число вложений")
    xlSheet.Range("A1:J1").Font.Bold = True

    row = 2
    count = 0

    Set appt = filtered.GetFirst
    Do While Not appt Is Nothing
        If TypeName(appt) = "AppointmentItem" Then
            If appt.Start >= startDate And appt.Start < DateAdd("d", 1, endDate) Then
                subjectText = NzStr(appt.Subject)
                If Not SubjectHasUppercaseDRWord(subjectText) Then
                    If Not SubjectMatchesExclusions(subjectText, exclusions) Then
                        If AppointmentHasAttendees(appt) Then
                            CollectAttachmentInfo appt, attachNames, attachCount
                            meetingLink = ExtractMeetingLinks(appt)
                            xlSheet.Cells(row, 1).Value = subjectText
                            xlSheet.Cells(row, 2).Value = Format$(appt.Start, "yyyy-mm-dd hh:nn")
                            xlSheet.Cells(row, 3).Value = NzStr(appt.Location)
                            xlSheet.Cells(row, 4).Value = meetingLink
                            xlSheet.Cells(row, 5).Value = FormatDurationMinutes(CLng(appt.Duration))
                            xlSheet.Cells(row, 6).Value = AppointmentOrganizer(appt)
                            xlSheet.Cells(row, 7).Value = AppointmentAttendeesByType(appt, olTo)
                            xlSheet.Cells(row, 8).Value = AppointmentAttendeesByType(appt, olCC)
                            xlSheet.Cells(row, 9).Value = attachNames
                            xlSheet.Cells(row, 10).Value = attachCount
                            row = row + 1
                            count = count + 1
                        End If
                    End If
                End If
            End If
        End If
        Set appt = filtered.GetNext
    Loop

    xlSheet.Columns("A:J").AutoFit

    savePath = BuildExportSavePath()
    On Error Resume Next
    xlBook.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        MsgBox "Выгрузка в Excel готова, но сохранить не удалось:" & vbCrLf & _
               savePath & vbCrLf & Err.Description, vbExclamation, "Сохранение"
        Err.Clear
    Else
        MsgBox "Календарь: " & calendarFolder.FolderPath & vbCrLf & _
               "Период: " & Format$(startDate, "dd.mm.yyyy") & " — " & _
               Format$(endDate, "dd.mm.yyyy") & vbCrLf & _
               "Встреч: " & CStr(count) & vbCrLf & _
               "Файл: " & savePath, vbInformation, "Выгрузка завершена"
    End If
    On Error GoTo 0

    xlApp.DisplayAlerts = True
End Sub

' --- Исключения тем ------------------------------------------------------

' False = пользователь отменил выгрузку. True = можно продолжать (список может быть пустым).
Private Function AskSubjectExclusions(ByRef exclusions() As String) As Boolean
    Dim raw As String
    Dim useSuggested As VbMsgBoxResult

    Erase exclusions

    useSuggested = MsgBox( _
        "Исключить типовые темы из выгрузки?" & vbCrLf & vbCrLf & _
        "Предложения: обед, ужин, lunch, dinner, focus time, block," & vbCrLf & _
        "занят, фокус, out of office, ooo, нн" & vbCrLf & vbCrLf & _
        "Да — подставить предложения в строку (можно править и дописывать)" & vbCrLf & _
        "Нет — пустая строка только для своих слов" & vbCrLf & _
        "Отмена — прервать выгрузку", _
        vbYesNoCancel Or vbQuestion, "Исключения из поиска")

    If useSuggested = vbCancel Then
        AskSubjectExclusions = False
        Exit Function
    End If

    If useSuggested = vbYes Then
        raw = DEFAULT_EXCLUSIONS
    Else
        raw = ""
    End If

    raw = InputBox( _
        "Слова/фразы для исключения (ищутся в теме встречи)." & vbCrLf & _
        "Несколько значений — через ; или ," & vbCrLf & _
        "Пример: обед;синк;1:1" & vbCrLf & vbCrLf & _
        "Можно удалить предложения, дописать свои или оставить пустым." & vbCrLf & _
        "Пусто = без фильтра по словам (кроме «ДР»).", _
        "Исключения из поиска", _
        raw)

    exclusions = ParseExclusionList(raw)
    AskSubjectExclusions = True
End Function

Private Function ParseExclusionList(ByVal raw As String) As String()
    Dim normalized As String
    Dim parts() As String
    Dim i As Long
    Dim token As String
    Dim tmp() As String
    Dim n As Long

    normalized = Replace(raw, ",", ";")
    parts = Split(normalized, ";")
    ReDim tmp(0 To UBound(parts))
    n = -1

    For i = LBound(parts) To UBound(parts)
        token = Trim$(parts(i))
        If Len(token) > 0 Then
            n = n + 1
            tmp(n) = token
        End If
    Next i

    If n < 0 Then
        Erase tmp
        ParseExclusionList = tmp
        Exit Function
    End If

    ReDim Preserve tmp(0 To n)
    ParseExclusionList = tmp
End Function

Private Function ExclusionCount(ByRef exclusions() As String) As Long
    On Error GoTo EmptyArr
    ExclusionCount = UBound(exclusions) - LBound(exclusions) + 1
    Exit Function
EmptyArr:
    ExclusionCount = 0
End Function

Private Function SubjectMatchesExclusions( _
    ByVal subject As String, _
    ByRef exclusions() As String) As Boolean

    Dim i As Long

    If ExclusionCount(exclusions) = 0 Then
        SubjectMatchesExclusions = False
        Exit Function
    End If

    For i = LBound(exclusions) To UBound(exclusions)
        If Len(exclusions(i)) > 0 Then
            If InStr(1, subject, exclusions(i), vbTextCompare) > 0 Then
                SubjectMatchesExclusions = True
                Exit Function
            End If
        End If
    Next i

    SubjectMatchesExclusions = False
End Function

' --- Прочее --------------------------------------------------------------

Private Function BuildExportSavePath() As String
    Dim folder As String
    Dim baseName As String
    Dim fullPath As String
    Dim sh As Object

    On Error Resume Next
    Set sh = CreateObject("WScript.Shell")
    If Not sh Is Nothing Then
        folder = sh.SpecialFolders("MyPictures")
    End If
    On Error GoTo 0

    If Len(folder) = 0 Then
        folder = Environ$("USERPROFILE") & "\Pictures"
    End If

    baseName = "calendar-export-" & Format$(Date, "yyyymmdd")
    fullPath = folder & "\" & baseName & ".xlsx"

    If Len(Dir$(fullPath)) > 0 Then
        fullPath = folder & "\" & baseName & "-" & Format$(Now, "hhnnss") & ".xlsx"
    End If

    BuildExportSavePath = fullPath
End Function

Private Function FormatDurationMinutes(ByVal minutes As Long) As String
    Dim hours As Long
    Dim mins As Long

    If minutes < 0 Then minutes = 0
    hours = minutes \ 60
    mins = minutes Mod 60
    FormatDurationMinutes = CStr(hours) & ":" & Right$("0" & CStr(mins), 2)
End Function

Private Function AppointmentHasAttendees(ByVal appt As Outlook.AppointmentItem) As Boolean
    Dim recipient As Outlook.Recipient

    On Error Resume Next
    If appt.MeetingStatus = olNonMeeting Then
        AppointmentHasAttendees = False
        On Error GoTo 0
        Exit Function
    End If

    For Each recipient In appt.Recipients
        If recipient.Type <> olOriginator Then
            AppointmentHasAttendees = True
            On Error GoTo 0
            Exit Function
        End If
    Next recipient
    On Error GoTo 0

    AppointmentHasAttendees = False
End Function

Private Function IsOutlook2016Family() As Boolean
    Dim ver As String
    Dim major As Long

    On Error Resume Next
    ver = Application.Version
    major = CLng(Val(ver))
    On Error GoTo 0

    IsOutlook2016Family = (major = 16)
End Function

Private Function PickCalendarFolder() As Outlook.MAPIFolder
    Dim folder As Outlook.MAPIFolder

    Set folder = Application.Session.PickFolder
    If folder Is Nothing Then
        Set PickCalendarFolder = Nothing
        Exit Function
    End If

    If folder.DefaultItemType <> olAppointmentItem Then
        MsgBox "Выберите папку календаря.", vbExclamation
        Set PickCalendarFolder = Nothing
        Exit Function
    End If

    Set PickCalendarFolder = folder
End Function

Private Function AskDateRange(ByRef startDate As Date, ByRef endDate As Date) As Boolean
    Dim startText As String
    Dim endText As String

    startText = InputBox( _
        "Начало периода (ДД.ММ.ГГГГ):", _
        "Диапазон дат", _
        Format$(DateSerial(Year(Date), Month(Date), 1), "dd.mm.yyyy"))
    If Len(Trim$(startText)) = 0 Then
        AskDateRange = False
        Exit Function
    End If

    endText = InputBox( _
        "Конец периода (ДД.ММ.ГГГГ):", _
        "Диапазон дат", _
        Format$(DateSerial(Year(Date), Month(Date) + 1, 0), "dd.mm.yyyy"))
    If Len(Trim$(endText)) = 0 Then
        AskDateRange = False
        Exit Function
    End If

    If Not TryParseDate(startText, startDate) Then
        MsgBox "Некорректная дата начала.", vbExclamation
        AskDateRange = False
        Exit Function
    End If

    If Not TryParseDate(endText, endDate) Then
        MsgBox "Некорректная дата конца.", vbExclamation
        AskDateRange = False
        Exit Function
    End If

    If endDate < startDate Then
        MsgBox "Конец периода раньше начала.", vbExclamation
        AskDateRange = False
        Exit Function
    End If

    startDate = DateSerial(Year(startDate), Month(startDate), Day(startDate))
    endDate = DateSerial(Year(endDate), Month(endDate), Day(endDate))

    AskDateRange = True
End Function

Private Function TryParseDate(ByVal text As String, ByRef result As Date) As Boolean
    On Error GoTo Fail
    result = CDate(Trim$(text))
    TryParseDate = True
    Exit Function
Fail:
    TryParseDate = False
End Function

Private Function QuoteOutlookDate(ByVal d As Date) As String
    QuoteOutlookDate = "'" & CStr(Month(d)) & "/" & CStr(Day(d)) & "/" & _
                       CStr(Year(d)) & " 12:00 AM'"
End Function

Private Function NzStr(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        NzStr = ""
    Else
        NzStr = CStr(value)
    End If
End Function

Private Function SubjectHasUppercaseDRWord(ByVal subject As String) As Boolean
    Dim j As Long
    Dim ch As String
    Dim buf As String
    Dim parts() As String
    Dim i As Long
    Dim token As String

    buf = vbNullString
    For j = 1 To Len(subject)
        ch = Mid$(subject, j, 1)
        Select Case ch
            Case " ", vbTab, vbCr, vbLf, ",", ";", ":", ".", "!", "?", _
                 "(", ")", "[", "]", "{", "}", "-", "/", "\", "|", """", "'"
                If Len(buf) > 0 Then
                    If Right$(buf, 1) <> " " Then buf = buf & " "
                End If
            Case Else
                buf = buf & ch
        End Select
    Next j

    buf = Trim$(buf)
    If Len(buf) = 0 Then
        SubjectHasUppercaseDRWord = False
        Exit Function
    End If

    parts = Split(buf, " ")
    For i = LBound(parts) To UBound(parts)
        token = Trim$(parts(i))
        If Len(token) > 0 Then
            If StrComp(token, "ДР", vbBinaryCompare) = 0 Then
                SubjectHasUppercaseDRWord = True
                Exit Function
            End If
        End If
    Next i

    SubjectHasUppercaseDRWord = False
End Function

Private Function AppointmentOrganizer(ByVal appt As Outlook.AppointmentItem) As String
    Dim organizer As Outlook.AddressEntry

    On Error Resume Next
    AppointmentOrganizer = Trim$(appt.Organizer)
    If Len(AppointmentOrganizer) = 0 Then
        Set organizer = appt.GetOrganizer
        If Not organizer Is Nothing Then
            AppointmentOrganizer = Trim$(organizer.Name)
        End If
    End If
    On Error GoTo 0
End Function

Private Function AppointmentAttendeesByType( _
    ByVal appt As Outlook.AppointmentItem, _
    ByVal recipientType As OlMailRecipientType) As String

    Dim recipient As Outlook.Recipient
    Dim names As String
    Dim person As String

    On Error Resume Next
    For Each recipient In appt.Recipients
        If recipient.Type = recipientType Then
            person = Trim$(recipient.Name)
            If Len(person) = 0 Then person = Trim$(recipient.Address)
            If Len(person) > 0 Then
                If Len(names) > 0 Then names = names & "; "
                names = names & person
            End If
        End If
    Next recipient
    On Error GoTo 0

    AppointmentAttendeesByType = names
End Function

' Только имена и количество — тело письма/встречи в Excel не пишем.
Private Sub CollectAttachmentInfo( _
    ByVal appt As Outlook.AppointmentItem, _
    ByRef namesOut As String, _
    ByRef countOut As Long)

    Dim attachment As Outlook.Attachment
    Dim fileName As String
    Dim names As String
    Dim n As Long

    names = ""
    n = 0

    On Error Resume Next
    For Each attachment In appt.Attachments
        fileName = Trim$(attachment.FileName)
        If Len(fileName) > 0 Then
            If Len(names) > 0 Then names = names & "; "
            names = names & fileName
            n = n + 1
        End If
    Next attachment
    On Error GoTo 0

    namesOut = names
    countOut = n
End Sub

' Ссылки Teams / Яндекс Телемост / ВКС Росатом — целиком, с путём и id
' (например https://ivavks.rosatom.ru/4321432...). Тело в Excel не пишем.
Private Function ExtractMeetingLinks(ByVal appt As Outlook.AppointmentItem) As String
    Dim haystack As String
    Dim bodyText As String

    haystack = NzStr(appt.Location)

    On Error Resume Next
    bodyText = appt.Body
    On Error GoTo 0
    If Len(bodyText) > 0 Then
        haystack = haystack & vbLf & bodyText
    End If

    ExtractMeetingLinks = FindAllMeetingLinks(haystack)
End Function

Private Function FindAllMeetingLinks(ByVal haystack As String) As String
    Dim hosts As Variant
    Dim h As Long
    Dim searchFrom As Long
    Dim hostPos As Long
    Dim url As String
    Dim found As String

    If Len(haystack) = 0 Then
        FindAllMeetingLinks = ""
        Exit Function
    End If

    ' Ищем по хостам: забираем полный URL включая /id?query=...
    hosts = Array( _
        "teams.microsoft.com", _
        "teams.live.com", _
        "telemost.yandex.ru", _
        "telemost.yandex.com", _
        "yandex.ru/telemost", _
        "ivavks.rosatom.ru")

    found = ""
    For h = LBound(hosts) To UBound(hosts)
        searchFrom = 1
        Do
            hostPos = InStr(searchFrom, haystack, CStr(hosts(h)), vbTextCompare)
            If hostPos = 0 Then Exit Do

            url = ExtractUrlAtHost(haystack, hostPos, CStr(hosts(h)))
            If Len(url) > 0 Then
                If InStr(1, found, url, vbTextCompare) = 0 Then
                    If Len(found) > 0 Then found = found & "; "
                    found = found & url
                End If
            End If
            searchFrom = hostPos + Len(CStr(hosts(h)))
        Loop
    Next h

    FindAllMeetingLinks = found
End Function

' От позиции хоста назад до http/https (если есть) и вперёд до конца URL — с путём и id.
Private Function ExtractUrlAtHost( _
    ByVal text As String, _
    ByVal hostPos As Long, _
    ByVal host As String) As String

    Dim startPos As Long
    Dim endPos As Long
    Dim i As Long
    Dim ch As String
    Dim url As String
    Dim prefix As String

    ' Начало: схема http(s):// или // или сразу хост
    startPos = hostPos
    If hostPos > 8 Then
        prefix = LCase$(Mid$(text, hostPos - 8, 8))
        If Right$(prefix, 8) = "https://" Then
            startPos = hostPos - 8
        End If
    End If
    If startPos = hostPos And hostPos > 7 Then
        prefix = LCase$(Mid$(text, hostPos - 7, 7))
        If Right$(prefix, 7) = "http://" Then
            startPos = hostPos - 7
        End If
    End If
    If startPos = hostPos And hostPos > 2 Then
        If Mid$(text, hostPos - 2, 2) = "//" Then
            startPos = hostPos - 2
        End If
    End If

    endPos = hostPos + Len(host)
    Do While endPos <= Len(text)
        ch = Mid$(text, endPos, 1)
        ' URL продолжается: путь, id, query, fragment
        Select Case True
            Case ch >= "A" And ch <= "Z"
            Case ch >= "a" And ch <= "z"
            Case ch >= "0" And ch <= "9"
            Case InStr(":/?#[]@!$&'()*+,;=-._~%", ch) > 0
            Case Else
                Exit Do
        End Select
        endPos = endPos + 1
    Loop

    url = Mid$(text, startPos, endPos - startPos)
    url = TrimTrailingUrlJunk(url)

    ' Если схемы не было — добавим https:// для кликабельности в Excel
    If Len(url) > 0 Then
        If LCase$(Left$(url, 7)) <> "http://" And LCase$(Left$(url, 8)) <> "https://" Then
            If Left$(url, 2) = "//" Then
                url = "https:" & url
            Else
                url = "https://" & url
            End If
        End If
    End If

    ExtractUrlAtHost = url
End Function

Private Function TrimTrailingUrlJunk(ByVal url As String) As String
    Dim lastChar As String

    Do While Len(url) > 0
        lastChar = Right$(url, 1)
        Select Case lastChar
            Case ".", ",", ";", "!", "?", ")", "]", "}", """", "'", ">"
                url = Left$(url, Len(url) - 1)
            Case Else
                Exit Do
        End Select
    Loop

    TrimTrailingUrlJunk = url
End Function
