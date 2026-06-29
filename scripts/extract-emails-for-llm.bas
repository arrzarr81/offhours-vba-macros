Attribute VB_Name = "EmailsForLLM"
' ============================================================================
'  Email Extractor for Local-LLM Categorization
'  Project: p5-outlook-offhours-emails (APS-llm-tools)
'
'  PURPOSE
'  -----------------------------------------------------------------
'  Exports an XLSX designed to be imported into a LOCAL LLM (GPT-OSS /
'  Gemma, running on the corporate device) so the model can answer:
'  "Would a 24-hour follow-the-sun APS support group resolve a meaningful
'  share of off-hours issues, or is that marginal?"
'
'  To answer that the LLM needs email CONTENT grouped into conversations,
'  so this macro - unlike find-offhours-emails.bas - DOES export the email
'  body. It also exports the native Outlook ConversationID (deterministic
'  thread grouping), the RFC Internet Message-ID (deterministic dedup), the
'  Inbound/Outbound direction, and deterministic team-involvement flags
'  (APS / ADM / DEV) matched from known addresses / distribution lists. The
'  LLM then judges substance, dead-ends, team, and APS-resolvability.
'
'  GOVERNANCE
'  -----------------------------------------------------------------
'  Email bodies are cleared to leave the mailbox ONLY inside this XLSX and
'  ONLY for local-LLM categorization. The file stays on the corporate
'  device. The other three macros in this repo remain metadata-only and
'  must stay that way - do not add body export to them.
'
'  READ-ONLY on Outlook. NO WRITE CAPABILITY.
'  -----------------------------------------------------------------
'  No .Save, no Categories assignment, no UnRead change, no Move / Delete /
'  Reply / Forward / Send anywhere in this module. It only READS:
'      ReceivedTime / SentOn, Subject, Body, Sender / SenderEmailAddress /
'      SenderName, Recipients, ConversationID, ConversationTopic, EntryID,
'      and the Internet Message-ID via PropertyAccessor.
'
'  How to use:
'    1. Outlook -> Alt+F11 -> File -> Import File... -> pick this .bas
'    2. Populate APS_ADDRESSES_CSV / ADM_ADDRESSES_CSV / DEV_ADDRESSES_CSV
'       and the usual CONFIG constants (mailbox, scan window).
'    3. Run macro: ExtractEmailsForLLM (F5)
'    4. Import the resulting XLSX into your local LLM together with
'       prompts/categorize-emails-followthesun.md
'
'  PRACTICAL: keep the scan window small (about one month) so the sheet
'  fits the local model's context window. Process longer periods in
'  monthly batches and sum the per-batch aggregates.
' ============================================================================

Option Explicit

' -------------------- CONFIG (edit these) --------------------
' Working hours (Paris local). Defines IsOffHours via the ORIGINAL Paris-gated
' rule (current coverage). "Off-hours" = the gap a follow-the-sun team fills.
Private Const WORK_START_HOUR As Integer = 9            ' inclusive, Paris local
Private Const WORK_END_HOUR As Integer = 18            ' exclusive, Paris local
Private Const TREAT_SATURDAY_AS_OFFHOURS As Boolean = True
Private Const TREAT_SUNDAY_AS_OFFHOURS As Boolean = True

' India coverage end on a Polish holiday (Paris local), 18:00 IST.
Private Const INDIA_END_SUMMER_MIN As Long = 14 * 60 + 30   ' 14:30 = 870
Private Const INDIA_END_WINTER_MIN As Long = 13 * 60 + 30   ' 13:30 = 810

' Holiday calendars (YYYY-MM-DD CSV). KEEP IN SYNC with the other macros.
Private Const HOLIDAYS_FR_CSV As String = _
    "2025-01-01,2025-04-21,2025-05-01,2025-05-08,2025-05-29," & _
    "2025-06-09,2025-07-14,2025-08-15,2025-11-01,2025-11-11,2025-12-25," & _
    "2026-01-01,2026-04-06,2026-05-01,2026-05-08,2026-05-14,2026-05-25," & _
    "2026-07-14,2026-08-15,2026-11-01,2026-11-11,2026-12-25"

Private Const HOLIDAYS_PL_CSV As String = _
    "2025-01-01,2025-01-06,2025-04-20,2025-04-21,2025-05-01,2025-05-03," & _
    "2025-06-08,2025-06-19,2025-08-15,2025-11-01,2025-11-11," & _
    "2025-12-24,2025-12-25,2025-12-26," & _
    "2026-01-01,2026-01-06,2026-04-05,2026-04-06,2026-05-01,2026-05-03," & _
    "2026-05-24,2026-06-04,2026-08-15,2026-11-01,2026-11-11," & _
    "2026-12-24,2026-12-25,2026-12-26"

Private Const HOLIDAYS_IN_MUMBAI_CSV As String = _
    "2025-01-26,2025-02-19,2025-03-14,2025-03-31,2025-04-10,2025-04-14," & _
    "2025-04-18,2025-05-01,2025-05-12,2025-06-07,2025-07-06,2025-08-15," & _
    "2025-08-27,2025-10-02,2025-10-21,2025-10-22,2025-11-05,2025-12-25," & _
    "2026-01-26,2026-03-04,2026-03-19,2026-04-03,2026-04-14,2026-05-01," & _
    "2026-08-15,2026-09-14,2026-10-02,2026-10-20,2026-11-08,2026-11-09," & _
    "2026-12-25"

' Team rosters - the deterministic half of team detection. Comma-separated.
' Each token is matched (case-insensitive substring) against From + To + Cc.
' Tokens may be full SMTP addresses ("aps-support@company.com"), domain
' fragments ("@aps."), or distribution-list display names ("APS Production").
' Leave "" to disable a team's deterministic match (LLM still infers from text).
Private Const APS_ADDRESSES_CSV As String = ""
Private Const ADM_ADDRESSES_CSV As String = ""
Private Const DEV_ADDRESSES_CSV As String = ""

' Mailbox to scan. "" = your own; display name or SMTP for a shared mailbox.
Private Const SHARED_MAILBOX As String = ""

' Primary (inbound) folder, plus optional Sent Items pass for outbound replies.
Private Const FOLDER_NAME As String = "Inbox"
Private Const INCLUDE_SUBFOLDERS As Boolean = False
Private Const SCAN_SENT_ALSO As Boolean = True    ' add Sent Items so threads include support replies

' Scan window (pick ONE pattern; see find-offhours-emails.bas for details).
Private Const SCAN_FROM_DATE As String = ""
Private Const SCAN_TO_DATE As String = ""
Private Const SCAN_LAST_N_DAYS As Long = 30

' Body handling.
Private Const MAX_BODY_CHARS As Long = 6000        ' Excel cell hard limit is 32767
Private Const STRIP_QUOTED_HISTORY As Boolean = True   ' cut quoted reply chains / signatures noise

' Output.
Private Const OUTPUT_DIR As String = ""            ' "" = <Documents>\OffHoursEmails\
Private Const ARCHIVE_FILES As Boolean = True
Private Const FILENAME_PREFIX As String = "emails-for-llm"

Private Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
Private Const COL_COUNT As Long = 25
' -------------------------------------------------------------


Public Sub ExtractEmailsForLLM()
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    Dim outPath As String
    outPath = ResolveOutputPath()
    If Len(outPath) = 0 Then Exit Sub

    ' Resolve scan window.
    Dim fromDt As Date, toDt As Date, windowDesc As String
    If Not ResolveWindow(fromDt, toDt, windowDesc) Then Exit Sub

    ' Build folder list with a parallel Direction list. Inbox tree = Inbound,
    ' Sent Items tree = Outbound.
    Dim folderList As Collection, pathList As Collection, dirList As Collection
    Set folderList = New Collection
    Set pathList = New Collection
    Set dirList = New Collection

    Dim inboxFolder As Outlook.folder
    Set inboxFolder = ResolveFolderByName(ns, FOLDER_NAME)
    If inboxFolder Is Nothing Then
        MsgBox "Folder not found. Mailbox=""" & SHARED_MAILBOX & """, Folder=""" & FOLDER_NAME & """", _
               vbExclamation, "Emails for LLM"
        Exit Sub
    End If
    Dim beforeCount As Long
    beforeCount = folderList.Count
    CollectFolders inboxFolder, "", folderList, pathList, INCLUDE_SUBFOLDERS
    AddDirection dirList, folderList.Count - beforeCount, "Inbound"

    If SCAN_SENT_ALSO Then
        Dim sentFolder As Outlook.folder
        Set sentFolder = ResolveFolderByName(ns, "Sent Items")
        If Not sentFolder Is Nothing Then
            beforeCount = folderList.Count
            CollectFolders sentFolder, "", folderList, pathList, INCLUDE_SUBFOLDERS
            AddDirection dirList, folderList.Count - beforeCount, "Outbound"
        End If
    End If

    ' Upper-bound row count to size the buffer.
    Dim maxRows As Long, fIdx As Long
    maxRows = 0
    For fIdx = 1 To folderList.Count
        maxRows = maxRows + folderList(fIdx).items.Count
    Next fIdx
    If maxRows < 1 Then maxRows = 1

    Dim data() As Variant
    ReDim data(1 To maxRows, 1 To COL_COUNT)
    Dim row As Long
    row = 0

    ' Dedup + stats accumulators.
    Dim seenMsgId As Object, threadSet As Object
    Set seenMsgId = CreateObject("Scripting.Dictionary")
    Set threadSet = CreateObject("Scripting.Dictionary")
    Dim total As Long, offHoursCount As Long, dupCount As Long
    Dim inboundCount As Long, outboundCount As Long
    Dim apsRows As Long, admRows As Long, devRows As Long

    Dim itm As Object
    For fIdx = 1 To folderList.Count
        Dim curFolder As Outlook.folder, curPath As String, curDir As String
        Set curFolder = folderList(fIdx)
        curPath = pathList(fIdx)
        curDir = dirList(fIdx)

        Dim dateField As String
        If curDir = "Outbound" Then dateField = "[SentOn]" Else dateField = "[ReceivedTime]"
        Dim restriction As String
        restriction = BuildRestriction(fromDt, toDt, dateField)

        Dim scan As Outlook.items
        Set scan = curFolder.items
        If Len(restriction) > 0 Then
            On Error Resume Next          ' a folder may not support the chosen field
            Set scan = scan.Restrict(restriction)
            On Error GoTo 0
        End If

        For Each itm In scan
            If TypeOf itm Is Outlook.MailItem Then
                Dim m As Outlook.MailItem
                Set m = itm
                total = total + 1

                ' When did it happen (received for inbound, sent for outbound)
                Dim whenDt As Date
                whenDt = SafeWhen(m, curDir)

                ' Timing block (original Paris-gated off-hours rule)
                Dim hFR As Boolean, hPL As Boolean, hIN As Boolean, isWknd As Boolean
                Dim hourLocal As Integer, minuteOfDay As Long, outsideOffice As Boolean, isOff As Boolean
                hFR = IsHolidayFR(whenDt)
                hPL = IsHolidayPL(whenDt)
                hIN = IsHolidayIN(whenDt)
                isWknd = IsWeekendDay(whenDt)
                hourLocal = Hour(whenDt)
                minuteOfDay = hourLocal * 60 + Minute(whenDt)
                outsideOffice = (minuteOfDay < WORK_START_HOUR * 60 Or minuteOfDay >= WORK_END_HOUR * 60)
                Dim indiaEndMin As Long
                If IsSummerTime(whenDt) Then indiaEndMin = INDIA_END_SUMMER_MIN Else indiaEndMin = INDIA_END_WINTER_MIN
                isOff = (isWknd And TREAT_SATURDAY_AS_OFFHOURS And Weekday(whenDt, vbMonday) = 6) Or _
                        (isWknd And TREAT_SUNDAY_AS_OFFHOURS And Weekday(whenDt, vbMonday) = 7) Or _
                        outsideOffice Or _
                        (hPL And (hIN Or minuteOfDay >= indiaEndMin))

                ' Identity / threading
                Dim threadId As String, topic As String, msgId As String
                threadId = SafeConversationId(m)
                topic = SafeStr(m.ConversationTopic)
                msgId = SafeMessageId(m)

                Dim isDup As Boolean
                isDup = False
                If Len(msgId) > 0 Then
                    If seenMsgId.Exists(msgId) Then
                        isDup = True
                        dupCount = dupCount + 1
                    Else
                        seenMsgId.Add msgId, True
                    End If
                End If
                If Len(threadId) > 0 Then
                    If Not threadSet.Exists(threadId) Then threadSet.Add threadId, True
                End If

                ' People / teams
                Dim fromAddr As String, toAddr As String, ccAddr As String
                fromAddr = SafeSenderAddress(m)
                toAddr = RecipientAddresses(m, olTo)
                ccAddr = RecipientAddresses(m, olCC)
                Dim blob As String
                blob = LCase(fromAddr & ";" & toAddr & ";" & ccAddr)
                Dim hasAPS As Boolean, hasADM As Boolean, hasDEV As Boolean
                hasAPS = ContainsAnyToken(blob, APS_ADDRESSES_CSV)
                hasADM = ContainsAnyToken(blob, ADM_ADDRESSES_CSV)
                hasDEV = ContainsAnyToken(blob, DEV_ADDRESSES_CSV)
                If hasAPS Then apsRows = apsRows + 1
                If hasADM Then admRows = admRows + 1
                If hasDEV Then devRows = devRows + 1

                Dim fromTeam As String
                fromTeam = TeamOf(LCase(fromAddr))

                ' Body
                Dim bodyClean As String, wasTrunc As Boolean
                bodyClean = CleanBody(SafeStr(m.Body), wasTrunc)

                If isOff Then offHoursCount = offHoursCount + 1
                If curDir = "Inbound" Then inboundCount = inboundCount + 1 Else outboundCount = outboundCount + 1

                row = row + 1
                data(row, 1) = threadId
                data(row, 2) = topic
                data(row, 3) = msgId
                data(row, 4) = isDup
                data(row, 5) = curDir
                data(row, 6) = Format(whenDt, "yyyy-mm-dd hh:nn:ss")
                data(row, 7) = hourLocal
                data(row, 8) = WeekdayName(Weekday(whenDt, vbMonday), True)
                data(row, 9) = isWknd
                data(row, 10) = hFR
                data(row, 11) = hPL
                data(row, 12) = hIN
                data(row, 13) = isOff
                data(row, 14) = fromAddr
                data(row, 15) = toAddr
                data(row, 16) = ccAddr
                data(row, 17) = fromTeam
                data(row, 18) = hasAPS
                data(row, 19) = hasADM
                data(row, 20) = hasDEV
                data(row, 21) = SafeStr(m.Subject)
                data(row, 22) = bodyClean
                data(row, 23) = wasTrunc
                data(row, 24) = curPath
                data(row, 25) = m.EntryID
            End If
        Next itm
    Next fIdx

    If row = 0 Then
        MsgBox "No mails in the scan window. No file written.", vbInformation, "Emails for LLM"
        Exit Sub
    End If

    WriteWorkbook outPath, data, row, maxRows, windowDesc, FolderSummary(folderList.Count), _
                  total, threadSet.Count, offHoursCount, dupCount, inboundCount, outboundCount, _
                  apsRows, admRows, devRows
End Sub


' ----------------------- Workbook writing -----------------------

Private Sub WriteWorkbook(ByVal outPath As String, ByRef data() As Variant, ByVal row As Long, _
                          ByVal maxRows As Long, ByVal windowDesc As String, ByVal folderDesc As String, _
                          ByVal total As Long, ByVal threadCount As Long, ByVal offHoursCount As Long, _
                          ByVal dupCount As Long, ByVal inboundCount As Long, ByVal outboundCount As Long, _
                          ByVal apsRows As Long, ByVal admRows As Long, ByVal devRows As Long)
    Dim xl As Object
    On Error Resume Next
    Set xl = CreateObject("Excel.Application")
    On Error GoTo 0
    If xl Is Nothing Then
        MsgBox "Could not start Excel. The macro requires Excel to write the XLSX.", vbCritical, "Emails for LLM"
        Exit Sub
    End If
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False

    Dim wb As Object
    Set wb = xl.Workbooks.Add
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop

    ' --- Sheet 1: Messages ---
    Dim ws As Object
    Set ws = wb.Worksheets(1)
    ws.Name = "Messages"

    Dim hdrs As Variant
    hdrs = Array("ThreadId", "ConversationTopic", "MessageId", "IsDuplicateMsgId", "Direction", _
                 "ReceivedTime", "HourLocal", "Weekday", "IsWeekend", _
                 "IsHolidayFR", "IsHolidayPL", "IsHolidayIN", "IsOffHours", _
                 "From", "To", "Cc", "FromTeam", "HasAPS", "HasADM", "HasDEV", _
                 "Subject", "Body", "BodyTruncated", "FolderPath", "EntryID")
    Dim hi As Long
    For hi = 0 To UBound(hdrs)
        ws.Cells(1, hi + 1).Value = hdrs(hi)
    Next hi
    ws.Range(ws.Cells(1, 1), ws.Cells(1, COL_COUNT)).Font.Bold = True

    ws.Range(ws.Cells(2, 1), ws.Cells(maxRows + 1, COL_COUNT)).Value = data
    If row < maxRows Then
        ws.Range(ws.Cells(row + 2, 1), ws.Cells(maxRows + 1, COL_COUNT)).ClearContents
    End If

    ' Group conversations: sort by ThreadId then ReceivedTime (string format is
    ' chronological), so each thread is a contiguous, ordered block.
    ws.Range(ws.Cells(1, 1), ws.Cells(row + 1, COL_COUNT)).Sort _
        Key1:=ws.Range("A2"), Order1:=1, Key2:=ws.Range("F2"), Order2:=1, Header:=1

    ws.Range(ws.Cells(1, 1), ws.Cells(row + 1, COL_COUNT)).AutoFilter
    ws.Columns("A:U").AutoFit
    ws.Columns("V").ColumnWidth = 80          ' Body - keep readable, not auto-ballooned
    ws.Columns("V").WrapText = False

    ' --- Sheet 2: Stats ---
    Dim wsStats As Object
    Set wsStats = wb.Worksheets.Add(After:=ws)
    wsStats.Name = "Stats"
    wsStats.Cells(1, 1).Value = "Emails for LLM - Run Statistics"
    wsStats.Cells(1, 1).Font.Bold = True
    wsStats.Cells(1, 1).Font.Size = 12

    Dim r As Long
    r = 3
    AddStatRow wsStats, r, "Run timestamp", Format(Now, "yyyy-mm-dd hh:nn:ss")
    AddStatRow wsStats, r, "Mailbox", IIf(Len(SHARED_MAILBOX) = 0, "<your own>", SHARED_MAILBOX)
    AddStatRow wsStats, r, "Folders scanned", folderDesc & IIf(SCAN_SENT_ALSO, " (Inbox + Sent)", "")
    AddStatRow wsStats, r, "Scan window", windowDesc
    AddStatRow wsStats, r, "Mode", "READ-ONLY (no mailbox changes)"
    AddStatRow wsStats, r, "Off-hours rule", "original Paris-gated (current coverage gap)"
    AddStatRow wsStats, r, "Max body chars", MAX_BODY_CHARS
    AddStatRow wsStats, r, "Output file", outPath
    r = r + 1

    SectionHeader wsStats, r, "Volume"
    AddStatRow wsStats, r, "Messages exported", row
    AddStatRow wsStats, r, "Distinct threads (ConversationID)", threadCount
    AddStatRow wsStats, r, "Inbound", inboundCount
    AddStatRow wsStats, r, "Outbound (Sent)", outboundCount
    AddStatRow wsStats, r, "Off-hours messages", offHoursCount
    AddStatRow wsStats, r, "Exact duplicates flagged (Message-ID)", dupCount
    r = r + 1

    SectionHeader wsStats, r, "Deterministic team matches (rows touching each team)"
    AddStatRow wsStats, r, "APS", apsRows
    AddStatRow wsStats, r, "ADM", admRows
    AddStatRow wsStats, r, "DEV", devRows
    If Len(APS_ADDRESSES_CSV) = 0 And Len(ADM_ADDRESSES_CSV) = 0 And Len(DEV_ADDRESSES_CSV) = 0 Then
        r = r + 1
        wsStats.Cells(r, 1).Value = "(team rosters are empty - the LLM will infer teams from content only)"
        wsStats.Cells(r, 1).Font.Italic = True
        wsStats.Cells(r, 1).Font.Color = RGB(150, 150, 150)
    End If
    wsStats.Columns("A:B").AutoFit

    ws.Activate
    xl.ActiveWindow.SplitRow = 1
    xl.ActiveWindow.FreezePanes = True
    ws.Range("A2").Select

    wb.SaveAs outPath, 51    ' 51 = xlOpenXMLWorkbook
    xl.ScreenUpdating = True
    xl.DisplayAlerts = True
    xl.Visible = True

    MsgBox "Mailbox: " & IIf(Len(SHARED_MAILBOX) = 0, "<your own>", SHARED_MAILBOX) & vbCrLf & _
           "Folders: " & folderDesc & IIf(SCAN_SENT_ALSO, " (Inbox + Sent)", "") & vbCrLf & _
           "Window:  " & windowDesc & vbCrLf & _
           "Mode:    READ-ONLY (no mailbox changes)" & vbCrLf & vbCrLf & _
           "Exported " & row & " messages across " & threadCount & " threads." & vbCrLf & _
           "  Off-hours messages:  " & offHoursCount & vbCrLf & _
           "  Duplicates flagged:  " & dupCount & vbCrLf & _
           "  Inbound / Outbound:  " & inboundCount & " / " & outboundCount & vbCrLf & vbCrLf & _
           "Next: import this XLSX into your local LLM with" & vbCrLf & _
           "prompts/categorize-emails-followthesun.md" & vbCrLf & vbCrLf & _
           "File: " & outPath, vbInformation, "Emails for LLM"
End Sub


' ----------------------- Scan-window helpers -----------------------

Private Function ResolveWindow(ByRef fromDt As Date, ByRef toDt As Date, ByRef windowDesc As String) As Boolean
    If Len(SCAN_FROM_DATE) > 0 Then
        On Error Resume Next
        fromDt = CDate(SCAN_FROM_DATE)
        If Err.Number <> 0 Then
            MsgBox "SCAN_FROM_DATE could not be parsed: """ & SCAN_FROM_DATE & """ (use YYYY-MM-DD).", _
                   vbExclamation, "Emails for LLM"
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    ElseIf SCAN_LAST_N_DAYS > 0 Then
        fromDt = DateAdd("d", -SCAN_LAST_N_DAYS, Now)
    End If

    If Len(SCAN_TO_DATE) > 0 Then
        On Error Resume Next
        toDt = CDate(SCAN_TO_DATE) + TimeSerial(23, 59, 59)
        If Err.Number <> 0 Then
            MsgBox "SCAN_TO_DATE could not be parsed: """ & SCAN_TO_DATE & """ (use YYYY-MM-DD).", _
                   vbExclamation, "Emails for LLM"
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If

    If fromDt > 0 And toDt > 0 Then
        windowDesc = Format(fromDt, "yyyy-mm-dd") & " to " & Format(toDt, "yyyy-mm-dd")
    ElseIf fromDt > 0 Then
        windowDesc = "from " & Format(fromDt, "yyyy-mm-dd") & " to now"
    ElseIf toDt > 0 Then
        windowDesc = "everything up to " & Format(toDt, "yyyy-mm-dd")
    Else
        windowDesc = "whole folder (no date filter)"
    End If
    ResolveWindow = True
End Function

Private Function BuildRestriction(ByVal fromDt As Date, ByVal toDt As Date, ByVal field As String) As String
    Dim s As String
    If fromDt > 0 Then s = field & " >= '" & Format(fromDt, "ddddd h:nn AMPM") & "'"
    If toDt > 0 Then
        If Len(s) > 0 Then s = s & " AND "
        s = s & field & " <= '" & Format(toDt, "ddddd h:nn AMPM") & "'"
    End If
    BuildRestriction = s
End Function

Private Function FolderSummary(ByVal n As Long) As String
    If n <= 1 Then FolderSummary = FOLDER_NAME Else FolderSummary = FOLDER_NAME & " + " & (n - 1) & " more"
End Function

Private Sub AddDirection(ByRef dirColl As Collection, ByVal count As Long, ByVal value As String)
    Dim i As Long
    For i = 1 To count
        dirColl.Add value
    Next i
End Sub


' ----------------------- Field-read helpers -----------------------

' Received time for inbound, sent time for outbound; robust to either missing.
Private Function SafeWhen(ByVal m As Outlook.MailItem, ByVal dir As String) As Date
    On Error Resume Next
    Dim d As Date
    If dir = "Outbound" Then
        d = m.SentOn
        If d <= 0 Then d = m.ReceivedTime
    Else
        d = m.ReceivedTime
        If d <= 0 Then d = m.SentOn
    End If
    On Error GoTo 0
    SafeWhen = d
End Function

Private Function SafeConversationId(ByVal m As Outlook.MailItem) As String
    On Error Resume Next
    SafeConversationId = m.ConversationID
    On Error GoTo 0
End Function

Private Function SafeMessageId(ByVal m As Outlook.MailItem) As String
    On Error Resume Next
    SafeMessageId = m.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
    On Error GoTo 0
End Function

Private Function SafeStr(ByVal v As Variant) As String
    On Error Resume Next
    SafeStr = CStr(v)
    On Error GoTo 0
End Function

' SMTP/address list for one recipient type (olTo / olCC), ";"-joined.
Private Function RecipientAddresses(ByVal m As Outlook.MailItem, ByVal recipType As Long) As String
    On Error Resume Next
    Dim out As String, i As Long, rcp As Outlook.Recipient, a As String
    For i = 1 To m.Recipients.Count
        Set rcp = m.Recipients(i)
        If rcp.Type = recipType Then
            a = ""
            a = rcp.Address
            If InStr(a, "@") = 0 Then
                Dim smtp As String
                smtp = rcp.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                If Len(smtp) > 0 Then a = smtp
            End If
            If Len(a) = 0 Then a = rcp.Name
            If Len(a) > 0 Then
                If Len(out) > 0 Then out = out & "; "
                out = out & a
            End If
        End If
    Next i
    On Error GoTo 0
    RecipientAddresses = out
End Function

Private Function SafeSenderAddress(ByVal m As Outlook.MailItem) As String
    On Error Resume Next
    Dim addr As String
    If Not m.Sender Is Nothing Then
        addr = m.Sender.GetExchangeUser().PrimarySmtpAddress
        If Len(addr) = 0 Then addr = m.SenderEmailAddress
    Else
        addr = m.SenderEmailAddress
    End If
    If Len(addr) = 0 Then addr = m.SenderName
    On Error GoTo 0
    SafeSenderAddress = addr
End Function


' ----------------------- Body cleaning -----------------------

' Strip quoted reply history / common signature dividers, collapse blank runs,
' and truncate to MAX_BODY_CHARS. Sets wasTruncated if the body was cut.
Private Function CleanBody(ByVal raw As String, ByRef wasTruncated As Boolean) As String
    wasTruncated = False
    If Len(raw) = 0 Then Exit Function
    Dim s As String
    s = raw

    If STRIP_QUOTED_HISTORY Then
        Dim cut As Long
        cut = EarliestMarker(s)
        If cut > 0 Then s = Left(s, cut - 1)
        s = RemoveQuotedLines(s)
    End If

    s = CollapseBlankLines(s)
    s = Trim(s)

    If Len(s) > MAX_BODY_CHARS Then
        s = Left(s, MAX_BODY_CHARS)
        wasTruncated = True
    End If
    CleanBody = s
End Function

' Index (1-based) of the earliest quoted-history / divider marker, else 0.
Private Function EarliestMarker(ByVal s As String) As Long
    Dim markers As Variant, i As Long, p As Long, best As Long
    markers = Array("-----Original Message-----", _
                    "________________________________", _
                    vbCrLf & "From: ", _
                    vbLf & "From: ", _
                    "-----Wiadomo", _
                    "-------- Forwarded Message --------")
    best = 0
    For i = LBound(markers) To UBound(markers)
        p = InStr(1, s, markers(i), vbTextCompare)
        If p > 0 Then
            If best = 0 Or p < best Then best = p
        End If
    Next i
    EarliestMarker = best
End Function

Private Function RemoveQuotedLines(ByVal s As String) As String
    Dim lines() As String, i As Long, out As String, ln As String
    lines = Split(Replace(s, vbCrLf, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = lines(i)
        If Left(LTrim(ln), 1) <> ">" Then
            out = out & ln & vbLf
        End If
    Next i
    RemoveQuotedLines = out
End Function

' Collapse 3+ consecutive newlines down to a maximum of two.
Private Function CollapseBlankLines(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    Do While InStr(s, vbLf & vbLf & vbLf) > 0
        s = Replace(s, vbLf & vbLf & vbLf, vbLf & vbLf)
    Loop
    CollapseBlankLines = s
End Function


' ----------------------- Team matching -----------------------

' True if any comma token of csv appears (case-insensitive) inside haystack.
Private Function ContainsAnyToken(ByVal haystackLower As String, ByVal csv As String) As Boolean
    If Len(csv) = 0 Then Exit Function
    Dim parts() As String, i As Long, tok As String
    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        tok = LCase(Trim(parts(i)))
        If Len(tok) > 0 Then
            If InStr(haystackLower, tok) > 0 Then ContainsAnyToken = True : Exit Function
        End If
    Next i
End Function

' Classify a sender address into a team (APS/ADM/DEV) or "Other".
Private Function TeamOf(ByVal fromLower As String) As String
    If ContainsAnyToken(fromLower, APS_ADDRESSES_CSV) Then TeamOf = "APS" : Exit Function
    If ContainsAnyToken(fromLower, ADM_ADDRESSES_CSV) Then TeamOf = "ADM" : Exit Function
    If ContainsAnyToken(fromLower, DEV_ADDRESSES_CSV) Then TeamOf = "DEV" : Exit Function
    TeamOf = "Other"
End Function


' ----------------------- Folder resolution -----------------------

' Resolve a named folder in the configured mailbox (own / shared by SMTP / by name).
Private Function ResolveFolderByName(ByVal ns As Outlook.NameSpace, ByVal folderName As String) As Outlook.folder
    On Error Resume Next
    If Len(SHARED_MAILBOX) = 0 Then
        If StrComp(folderName, "Sent Items", vbTextCompare) = 0 Then
            Set ResolveFolderByName = ns.GetDefaultFolder(olFolderSentMail)
        ElseIf StrComp(folderName, "Inbox", vbTextCompare) = 0 Then
            Set ResolveFolderByName = ns.GetDefaultFolder(olFolderInbox)
        Else
            Set ResolveFolderByName = ns.GetDefaultFolder(olFolderInbox).folders(folderName)
        End If
        Exit Function
    End If

    If InStr(SHARED_MAILBOX, "@") > 0 Then
        Dim recip As Outlook.Recipient
        Set recip = ns.CreateRecipient(SHARED_MAILBOX)
        recip.Resolve
        If Not recip.Resolved Then Exit Function
        If StrComp(folderName, "Sent Items", vbTextCompare) = 0 Then
            Set ResolveFolderByName = ns.GetSharedDefaultFolder(recip, olFolderSentMail)
        Else
            Dim sharedRoot As Outlook.folder
            Set sharedRoot = ns.GetSharedDefaultFolder(recip, olFolderInbox)
            If StrComp(folderName, "Inbox", vbTextCompare) = 0 Then
                Set ResolveFolderByName = sharedRoot
            Else
                Set ResolveFolderByName = sharedRoot.folders(folderName)
            End If
        End If
        Exit Function
    End If

    Dim store As Outlook.folder
    Set store = ns.folders(SHARED_MAILBOX)
    If store Is Nothing Then Exit Function
    Set ResolveFolderByName = store.folders(folderName)
End Function

' Depth-first folder walk into parallel folder/path collections (read-only).
Private Sub CollectFolders(ByVal root As Outlook.folder, ByVal pathSoFar As String, _
                           ByRef folderColl As Collection, ByRef pathColl As Collection, _
                           ByVal includeSubfolders As Boolean)
    Dim rootPath As String
    If Len(pathSoFar) = 0 Then rootPath = root.Name Else rootPath = pathSoFar & "/" & root.Name
    folderColl.Add root
    pathColl.Add rootPath
    If Not includeSubfolders Then Exit Sub
    Dim child As Outlook.folder
    For Each child In root.folders
        CollectFolders child, rootPath, folderColl, pathColl, True
    Next child
End Sub


' ----------------------- Timing helpers -----------------------

Private Function IsWeekendDay(ByVal t As Date) As Boolean
    Dim dow As Integer
    dow = Weekday(t, vbMonday)
    IsWeekendDay = (dow = 6 Or dow = 7)
End Function

Private Function IsInDateList(ByVal t As Date, ByVal listCsv As String) As Boolean
    If Len(listCsv) = 0 Then Exit Function
    Dim key As String
    key = Format(t, "yyyy-mm-dd")
    Dim parts() As String, i As Long
    parts = Split(listCsv, ",")
    For i = LBound(parts) To UBound(parts)
        If Trim(parts(i)) = key Then IsInDateList = True : Exit Function
    Next i
End Function

Private Function IsHolidayFR(ByVal t As Date) As Boolean
    IsHolidayFR = IsInDateList(t, HOLIDAYS_FR_CSV)
End Function
Private Function IsHolidayPL(ByVal t As Date) As Boolean
    IsHolidayPL = IsInDateList(t, HOLIDAYS_PL_CSV)
End Function
Private Function IsHolidayIN(ByVal t As Date) As Boolean
    IsHolidayIN = IsInDateList(t, HOLIDAYS_IN_MUMBAI_CSV)
End Function

Private Function IsSummerTime(ByVal t As Date) As Boolean
    Dim y As Integer
    y = Year(t)
    IsSummerTime = (DateValue(t) >= LastSundayOfMonth(y, 3) And DateValue(t) < LastSundayOfMonth(y, 10))
End Function

Private Function LastSundayOfMonth(ByVal y As Integer, ByVal m As Integer) As Date
    Dim lastDay As Date, dow As Integer
    lastDay = DateSerial(y, m + 1, 0)
    dow = Weekday(lastDay, vbMonday)
    LastSundayOfMonth = lastDay - (dow Mod 7)
End Function


' ----------------------- Stats / output path helpers -----------------------

Private Sub AddStatRow(ByVal ws As Object, ByRef r As Long, ByVal label As String, ByVal value As Variant)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Value = value
    r = r + 1
End Sub

Private Sub SectionHeader(ByVal ws As Object, ByRef r As Long, ByVal title As String)
    ws.Cells(r, 1).Value = title
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
End Sub

Private Function ResolveOutputPath() As String
    Dim outDir As String
    If Len(OUTPUT_DIR) > 0 Then outDir = OUTPUT_DIR Else outDir = GetDocumentsPath() & "\OffHoursEmails"
    If Right(outDir, 1) <> "\" Then outDir = outDir & "\"

    On Error Resume Next
    If Len(Dir(outDir, vbDirectory)) = 0 Then MkDir outDir
    On Error GoTo 0
    If Len(Dir(outDir, vbDirectory)) = 0 Then
        MsgBox "Cannot create output directory:" & vbCrLf & outDir, vbCritical, "Emails for LLM"
        Exit Function
    End If

    If ARCHIVE_FILES Then
        Dim stamp As String
        stamp = Format(Now, "yyyy-mm-dd") & "_" & Format(Now, "hhnn")
        If Len(SHARED_MAILBOX) > 0 Then
            ResolveOutputPath = outDir & FILENAME_PREFIX & "-" & SanitizeForFilename(SHARED_MAILBOX) & "-" & stamp & ".xlsx"
        Else
            ResolveOutputPath = outDir & FILENAME_PREFIX & "-" & stamp & ".xlsx"
        End If
    Else
        ResolveOutputPath = outDir & FILENAME_PREFIX & ".xlsx"
    End If
End Function

Private Function SanitizeForFilename(ByVal s As String) As String
    Dim out As String, i As Long, c As String
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c Like "[A-Za-z0-9_-]" Then
            out = out & c
        ElseIf c = " " Then
            out = out & "_"
        End If
    Next i
    SanitizeForFilename = out
End Function

Private Function GetDocumentsPath() As String
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    If Not wsh Is Nothing Then GetDocumentsPath = wsh.SpecialFolders("MyDocuments")
    On Error GoTo 0
    If Len(GetDocumentsPath) = 0 Then GetDocumentsPath = Environ("USERPROFILE") & "\Documents"
End Function
