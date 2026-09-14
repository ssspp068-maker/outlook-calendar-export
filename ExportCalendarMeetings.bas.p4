pointmentItem) As String
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

Private Function ExtractUrlAtHost( _