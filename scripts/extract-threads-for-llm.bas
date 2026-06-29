Attribute VB_Name = "ThreadsForLLM"
' ============================================================================
'  Thread Extractor for Local-LLM Categorization  (v2 - thread level)
'  Project: p5-outlook-offhours-emails (APS-llm-tools)
'
'  RELATIONSHIP TO v1
'  -----------------------------------------------------------------
'  v1 = extract-emails-for-llm.bas: exports ONE ROW PER MESSAGE and lets the
'       LLM do everything in a single freeform pass. Kept unchanged.
'  v2 = THIS module: pre-aggregates to ONE ROW PER CONVERSATION, computes the
'       deterministic judgments (dedup, off-hours trigger, automated sender,
'       dead-end heuristic, recurrence) in VBA, and feeds the LLM a condensed
'       row per thread for a reproducible, two-pass (induce taxonomy -> apply)
'       classification whose aggregates are counted in Excel, not by the LLM.
'
'  Run BOTH on the same window and compare (see prompts/threadlevel-0-runbook.md).
'
'  PURPOSE
'  -----------------------------------------------------------------
'  Answer: "Would a 24-hour follow-the-sun APS team resolve a meaningful share
'  of off-hours issues, or is it marginal?" The headline metric the workbook
'  builds toward is:
'      genuine_issue AND offhours_trigger AND NOT dead_end
'                     AND aps_resolvable_alone AND time_sensitive
'  plus a recurrence view (N instances of one fixable root cause).
'
'  GOVERNANCE / SAFETY  (identical to v1)
'  -----------------------------------------------------------------
'  Email bodies leave the mailbox ONLY inside this XLSX, ONLY for local-LLM
'  categorization, and the file stays on the corporate device.
'  READ-ONLY on Outlook: no .Save / Categories= / UnRead= / Move / Delete /
'  Reply / Forward / Send on any mail item. Reads only ReceivedTime / SentOn,
'  Subject, Body, Sender*, Recipients, ConversationID/Topic, EntryID, and the
'  Internet Message-ID via PropertyAccessor.
'
'  How to use:
'    1. Outlook -> Alt+F11 -> File -> Import File... -> pick this .bas
'    2. Populate SUPPORT_ADDRESSES_CSV (your project inbox + APS sender
'       addresses - drives the Inbound/Outbound split), the team rosters,
'       AUTOMATED_SENDERS_CSV, and the mailbox/window constants.
'       Sent Items is NOT needed if replies are CC'd to the project inbox.
'    3. Run ExtractThreadsForLLM (F5).
'    4. Follow prompts/threadlevel-0-runbook.md (pass A -> Taxonomy,
'       pass B -> Classification, read Summary, calibrate on GoldSet).
' ============================================================================

Option Explicit

' -------------------- CONFIG (edit these) --------------------
' Off-hours uses the ORIGINAL Paris-gated rule (current coverage gap).
Private Const WORK_START_HOUR As Integer = 9
Private Const WORK_END_HOUR As Integer = 18
Private Const TREAT_SATURDAY_AS_OFFHOURS As Boolean = True
Private Const TREAT_SUNDAY_AS_OFFHOURS As Boolean = True
Private Const INDIA_END_SUMMER_MIN As Long = 14 * 60 + 30   ' 14:30
Private Const INDIA_END_WINTER_MIN As Long = 13 * 60 + 30   ' 13:30

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

' Team rosters (deterministic half of team detection). Matched case-insensitive
' against From + To + Cc. Tokens: SMTP addresses, "@domain." fragments, or DL
' display names. "" disables a team's deterministic match (LLM still infers).
Private Const APS_ADDRESSES_CSV As String = ""
Private Const ADM_ADDRESSES_CSV As String = ""
Private Const DEV_ADDRESSES_CSV As String = ""

' OUR side = the support desk. A message whose From matches is tagged Outbound
' (a support reply), EVEN when it sits in the Inbox as a CC'd copy; everything
' else is Inbound. This is how Direction is decided - by sender, not by folder -
' so you do NOT need to scan Sent Items as long as your team CCs the project
' inbox on every reply. Put the project/shared mailbox address + your APS sender
' addresses here (add ADM/DEV addresses too if you want their replies treated as
' outbound). The shared mailbox SMTP in SHARED_MAILBOX is matched automatically.
' MUST be populated for this Inbox-only model to find support replies.
Private Const SUPPORT_ADDRESSES_CSV As String = ""

' Automated/no-reply senders to flag as noise (matched against From).
Private Const AUTOMATED_SENDERS_CSV As String = _
    "no-reply,noreply,donotreply,do-not-reply,mailer-daemon,postmaster," & _
    "notification,notifications,alert,alerts,monitoring,nagios,zabbix," & _
    "dynatrace,splunk,jenkins,jira@,servicenow,automation"

' Mailbox + scope.
Private Const SHARED_MAILBOX As String = ""
Private Const FOLDER_NAME As String = "Inbox"
Private Const INCLUDE_SUBFOLDERS As Boolean = False
' Inbox-only by default: replies CC'd into the project inbox are already in the
' thread, and Direction is decided by sender (SUPPORT_ADDRESSES_CSV), so the
' Sent folder is redundant. Set True only for mailboxes that do NOT CC themselves.
Private Const SCAN_SENT_ALSO As Boolean = False

' Scan window (pick ONE pattern).
Private Const SCAN_FROM_DATE As String = ""
Private Const SCAN_TO_DATE As String = ""
Private Const SCAN_LAST_N_DAYS As Long = 30

' Body / condensation.
Private Const MAX_BODY_CHARS As Long = 6000        ' per-message Body on the Messages sheet
Private Const STRIP_QUOTED_HISTORY As Boolean = True
Private Const COND_FIRST_CHARS As Long = 1500      ' first inbound body in CondensedText
Private Const COND_LAST_CHARS As Long = 800        ' last support reply in CondensedText

' Thread heuristics.
Private Const DEADEND_MAX_CHARS As Long = 200       ' "thanks"-length cap for dead-end test
Private Const RECURRENCE_KEY As String = "subject"  ' "subject" or "subject+sender"
Private Const TAXONOMY_SAMPLE_SIZE As Long = 100    ' threads flagged SampleForTaxonomy (pass A)

' Output.
Private Const OUTPUT_DIR As String = ""
Private Const ARCHIVE_FILES As Boolean = True
Private Const FILENAME_PREFIX As String = "threads-for-llm"

Private Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"

' Per-message parallel arrays (1..msgCount).
Private mConv() As String, mTopic() As String, mMsgId() As String
Private mIsDup() As Boolean, mDir() As String, mWhen() As Date
Private mOff() As Boolean, mFrom() As String, mFromTeam() As String
Private mHasAPS() As Boolean, mHasADM() As Boolean, mHasDEV() As Boolean
Private mAuto() As Boolean, mSubject() As String, mBody() As String
Private mFolder() As String, mEntry() As String
' -------------------------------------------------------------


Public Sub ExtractThreadsForLLM()
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    Dim outPath As String
    outPath = ResolveOutputPath()
    If Len(outPath) = 0 Then Exit Sub

    Dim fromDt As Date, toDt As Date, windowDesc As String
    If Not ResolveWindow(fromDt, toDt, windowDesc) Then Exit Sub

    ' Folder list (Inbox tree = Inbound, Sent tree = Outbound).
    Dim folderList As Collection, pathList As Collection, dirList As Collection
    Set folderList = New Collection: Set pathList = New Collection: Set dirList = New Collection

    Dim inboxFolder As Outlook.folder
    Set inboxFolder = ResolveFolderByName(ns, FOLDER_NAME)
    If inboxFolder Is Nothing Then
        MsgBox "Folder not found. Mailbox=""" & SHARED_MAILBOX & """, Folder=""" & FOLDER_NAME & """", _
               vbExclamation, "Threads for LLM"
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

    Dim maxRows As Long, fIdx As Long
    maxRows = 0
    For fIdx = 1 To folderList.Count
        maxRows = maxRows + folderList(fIdx).items.Count
    Next fIdx
    If maxRows < 1 Then maxRows = 1

    ReDim mConv(1 To maxRows): ReDim mTopic(1 To maxRows): ReDim mMsgId(1 To maxRows)
    ReDim mIsDup(1 To maxRows): ReDim mDir(1 To maxRows): ReDim mWhen(1 To maxRows)
    ReDim mOff(1 To maxRows): ReDim mFrom(1 To maxRows): ReDim mFromTeam(1 To maxRows)
    ReDim mHasAPS(1 To maxRows): ReDim mHasADM(1 To maxRows): ReDim mHasDEV(1 To maxRows)
    ReDim mAuto(1 To maxRows): ReDim mSubject(1 To maxRows): ReDim mBody(1 To maxRows)
    ReDim mFolder(1 To maxRows): ReDim mEntry(1 To maxRows)

    Dim seenMsgId As Object
    Set seenMsgId = CreateObject("Scripting.Dictionary")
    Dim row As Long, total As Long, offHoursCount As Long, dupCount As Long
    Dim inboundCount As Long, outboundCount As Long
    row = 0

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
            On Error Resume Next
            Set scan = scan.Restrict(restriction)
            On Error GoTo 0
        End If

        For Each itm In scan
            If TypeOf itm Is Outlook.MailItem Then
                Dim m As Outlook.MailItem
                Set m = itm
                total = total + 1

                Dim conv As String
                conv = SafeConversationId(m)
                If Len(conv) = 0 Then conv = "NOCONV-" & m.EntryID

                Dim whenDt As Date
                whenDt = SafeWhen(m, curDir)

                Dim isOff As Boolean
                isOff = ComputeOffHours(whenDt)

                Dim fromAddr As String, toAddr As String, ccAddr As String, blob As String
                fromAddr = SafeSenderAddress(m)
                toAddr = RecipientAddresses(m, olTo)
                ccAddr = RecipientAddresses(m, olCC)
                blob = LCase(fromAddr & ";" & toAddr & ";" & ccAddr)

                ' Direction by SENDER, not folder: a support-side sender = Outbound
                ' reply (even as a CC'd copy in the Inbox); else Inbound.
                Dim msgDir As String
                If IsSupportSender(LCase(fromAddr)) Then msgDir = "Outbound" Else msgDir = "Inbound"

                Dim msgId As String, isDup As Boolean
                msgId = SafeMessageId(m)
                isDup = False
                If Len(msgId) > 0 Then
                    If seenMsgId.Exists(msgId) Then
                        isDup = True : dupCount = dupCount + 1
                    Else
                        seenMsgId.Add msgId, True
                    End If
                End If

                Dim wasTrunc As Boolean
                row = row + 1
                mConv(row) = conv
                mTopic(row) = SafeStr(m.ConversationTopic)
                mMsgId(row) = msgId
                mIsDup(row) = isDup
                mDir(row) = msgDir
                mWhen(row) = whenDt
                mOff(row) = isOff
                mFrom(row) = fromAddr
                mFromTeam(row) = TeamOf(LCase(fromAddr))
                mHasAPS(row) = ContainsAnyToken(blob, APS_ADDRESSES_CSV)
                mHasADM(row) = ContainsAnyToken(blob, ADM_ADDRESSES_CSV)
                mHasDEV(row) = ContainsAnyToken(blob, DEV_ADDRESSES_CSV)
                mAuto(row) = ContainsAnyToken(LCase(fromAddr), AUTOMATED_SENDERS_CSV)
                mSubject(row) = SafeStr(m.Subject)
                mBody(row) = CleanBody(SafeStr(m.Body), wasTrunc)
                mFolder(row) = curPath
                mEntry(row) = m.EntryID

                If isOff Then offHoursCount = offHoursCount + 1
                If msgDir = "Inbound" Then inboundCount = inboundCount + 1 Else outboundCount = outboundCount + 1
            End If
        Next itm
    Next fIdx

    If row = 0 Then
        MsgBox "No mails in the scan window. No file written.", vbInformation, "Threads for LLM"
        Exit Sub
    End If

    ' ---- Aggregate messages into threads ----
    Dim tDict As Object
    Set tDict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To row
        If Not tDict.Exists(mConv(i)) Then tDict.Add mConv(i), New Collection
        tDict(mConv(i)).Add i
    Next i

    Dim tCount As Long
    tCount = tDict.Count
    Dim tData() As Variant
    ReDim tData(1 To tCount, 1 To 20)

    ' Recurrence counting (first pass over threads to count keys).
    Dim recCount As Object, recId As Object
    Set recCount = CreateObject("Scripting.Dictionary")
    Set recId = CreateObject("Scripting.Dictionary")

    Dim tIdx As Long, k As Variant
    Dim recKeys() As String
    ReDim recKeys(1 To tCount)

    ' First loop: compute per-thread features + recurrence keys.
    tIdx = 0
    For Each k In tDict.Keys
        tIdx = tIdx + 1
        Dim rows As Collection
        Set rows = tDict(k)

        Dim firstInRow As Long, firstInTime As Date
        Dim firstAnyRow As Long, firstAnyTime As Date
        Dim lastRow As Long, lastTime As Date, lastDir As String
        Dim lastOutRow As Long, lastOutTime As Date
        Dim inC As Long, outC As Long
        Dim hAPS As Boolean, hADM As Boolean, hDEV As Boolean, anyInOff As Boolean
        firstInRow = 0: firstAnyRow = 0: lastRow = 0: lastOutRow = 0
        inC = 0: outC = 0

        Dim jj As Variant, idx As Long
        For Each jj In rows
            idx = jj
            If firstAnyRow = 0 Or mWhen(idx) < firstAnyTime Then firstAnyRow = idx : firstAnyTime = mWhen(idx)
            If lastRow = 0 Or mWhen(idx) > lastTime Then lastRow = idx : lastTime = mWhen(idx) : lastDir = mDir(idx)
            If mDir(idx) = "Inbound" Then
                inC = inC + 1
                If firstInRow = 0 Or mWhen(idx) < firstInTime Then firstInRow = idx : firstInTime = mWhen(idx)
                If mOff(idx) Then anyInOff = True
            Else
                outC = outC + 1
                If lastOutRow = 0 Or mWhen(idx) > lastOutTime Then lastOutRow = idx : lastOutTime = mWhen(idx)
            End If
            If mHasAPS(idx) Then hAPS = True
            If mHasADM(idx) Then hADM = True
            If mHasDEV(idx) Then hDEV = True
        Next jj

        Dim triggerRow As Long
        If firstInRow > 0 Then triggerRow = firstInRow Else triggerRow = firstAnyRow

        Dim offHoursTrigger As Boolean
        offHoursTrigger = (firstInRow > 0) And mOff(firstInRow)

        ' Dead-end: second loop over off-hours messages (needs lastOutTime).
        Dim offCnt As Long, nonActCnt As Long
        offCnt = 0: nonActCnt = 0
        For Each jj In rows
            idx = jj
            If mOff(idx) Then
                offCnt = offCnt + 1
                If IsNonActionable(idx, lastOutTime) Then nonActCnt = nonActCnt + 1
            End If
        Next jj
        Dim deadEnd As Boolean
        deadEnd = (offCnt > 0) And (offCnt = nonActCnt)

        Dim topicStr As String
        topicStr = mTopic(firstAnyRow)
        If Len(topicStr) = 0 Then topicStr = mSubject(firstAnyRow)
        Dim normSubj As String
        normSubj = NormalizeSubject(topicStr)

        Dim recKey As String
        If LCase(RECURRENCE_KEY) = "subject+sender" Then
            recKey = normSubj & "|" & LCase(mFrom(triggerRow))
        Else
            recKey = normSubj
        End If
        recKeys(tIdx) = recKey
        If recCount.Exists(recKey) Then
            recCount(recKey) = recCount(recKey) + 1
        Else
            recCount.Add recKey, 1
            recId.Add recKey, recId.Count + 1
        End If

        Dim condensed As String
        condensed = "INBOUND: " & Left(mBody(triggerRow), COND_FIRST_CHARS)
        If lastOutRow > 0 Then
            condensed = condensed & vbLf & "---" & vbLf & "SUPPORT REPLY: " & Left(mBody(lastOutRow), COND_LAST_CHARS)
        End If

        tData(tIdx, 1) = mConv(triggerRow)                       ' ThreadId
        tData(tIdx, 2) = topicStr                                ' ConversationTopic
        tData(tIdx, 3) = Format(firstAnyTime, "yyyy-mm-dd hh:nn:ss")   ' FirstInboundTime (earliest msg)
        tData(tIdx, 4) = Format(lastTime, "yyyy-mm-dd hh:nn:ss")  ' LastMessageTime
        tData(tIdx, 5) = lastDir                                 ' LastMessageDirection
        tData(tIdx, 6) = rows.Count                              ' MsgCount
        tData(tIdx, 7) = inC                                     ' InboundCount
        tData(tIdx, 8) = outC                                    ' OutboundCount
        tData(tIdx, 9) = offHoursTrigger                         ' OffHoursTrigger
        tData(tIdx, 10) = anyInOff                               ' AnyInboundOffHours
        tData(tIdx, 11) = mAuto(triggerRow)                      ' AutomatedSender
        tData(tIdx, 12) = deadEnd                                ' DeadEndHeuristic
        tData(tIdx, 13) = hAPS                                   ' HasAPS
        tData(tIdx, 14) = hADM                                   ' HasADM
        tData(tIdx, 15) = hDEV                                   ' HasDEV
        tData(tIdx, 16) = 0                                      ' RecurrenceGroupId (filled below)
        tData(tIdx, 17) = 0                                      ' RecurrenceCount (filled below)
        tData(tIdx, 18) = normSubj                               ' NormalizedSubject
        tData(tIdx, 19) = False                                  ' SampleForTaxonomy (filled below)
        tData(tIdx, 20) = condensed                              ' CondensedText
    Next k

    ' Second pass: fill recurrence id/count + taxonomy sample flag.
    Dim sampleEvery As Long, sampledSoFar As Long
    If TAXONOMY_SAMPLE_SIZE > 0 Then sampleEvery = tCount \ TAXONOMY_SAMPLE_SIZE
    If sampleEvery < 1 Then sampleEvery = 1
    For tIdx = 1 To tCount
        tData(tIdx, 16) = recId(recKeys(tIdx))
        tData(tIdx, 17) = recCount(recKeys(tIdx))
        If sampledSoFar < TAXONOMY_SAMPLE_SIZE And (tIdx Mod sampleEvery = 0) Then
            tData(tIdx, 19) = True
            sampledSoFar = sampledSoFar + 1
        End If
    Next tIdx

    WriteWorkbook outPath, tData, tCount, row, maxRows, windowDesc, FolderSummary(folderList.Count), _
                  total, offHoursCount, dupCount, inboundCount, outboundCount, recCount.Count
End Sub


' ----------------------- Workbook writing -----------------------

Private Sub WriteWorkbook(ByVal outPath As String, ByRef tData() As Variant, ByVal tCount As Long, _
                          ByVal msgCount As Long, ByVal maxRows As Long, ByVal windowDesc As String, _
                          ByVal folderDesc As String, ByVal total As Long, ByVal offHoursCount As Long, _
                          ByVal dupCount As Long, ByVal inboundCount As Long, ByVal outboundCount As Long, _
                          ByVal recurGroups As Long)
    Dim xl As Object
    On Error Resume Next
    Set xl = CreateObject("Excel.Application")
    On Error GoTo 0
    If xl Is Nothing Then
        MsgBox "Could not start Excel. The macro requires Excel to write the XLSX.", vbCritical, "Threads for LLM"
        Exit Sub
    End If
    xl.Visible = False: xl.DisplayAlerts = False: xl.ScreenUpdating = False

    Dim wb As Object
    Set wb = xl.Workbooks.Add
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop

    ' Sheet order: Threads, Taxonomy, Classification, Summary, GoldSet, Messages, Stats
    Dim wsThreads As Object
    Set wsThreads = wb.Worksheets(1)
    wsThreads.Name = "Threads"
    BuildThreadsSheet wsThreads, tData, tCount

    Dim wsTax As Object: Set wsTax = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    BuildTaxonomySheet wsTax
    Dim wsClass As Object: Set wsClass = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    BuildClassificationSheet wsClass
    Dim wsSum As Object: Set wsSum = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    BuildSummarySheet wsSum
    Dim wsGold As Object: Set wsGold = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    BuildGoldSetSheet wsGold
    Dim wsMsg As Object: Set wsMsg = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    BuildMessagesSheet wsMsg, msgCount, maxRows
    Dim wsStats As Object: Set wsStats = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    BuildStatsSheet wsStats, outPath, windowDesc, folderDesc, total, msgCount, tCount, _
                    offHoursCount, dupCount, inboundCount, outboundCount, recurGroups

    wsThreads.Activate
    xl.ActiveWindow.SplitRow = 1
    xl.ActiveWindow.FreezePanes = True
    wsThreads.Range("A2").Select

    wb.SaveAs outPath, 51
    xl.ScreenUpdating = True: xl.DisplayAlerts = True: xl.Visible = True

    MsgBox "Threads for LLM (v2)" & vbCrLf & _
           "Mailbox: " & IIf(Len(SHARED_MAILBOX) = 0, "<your own>", SHARED_MAILBOX) & vbCrLf & _
           "Window:  " & windowDesc & vbCrLf & _
           "Mode:    READ-ONLY" & vbCrLf & vbCrLf & _
           "Messages: " & msgCount & "   Threads: " & tCount & vbCrLf & _
           "Off-hours messages: " & offHoursCount & "   Duplicates: " & dupCount & vbCrLf & vbCrLf & _
           "Next: prompts/threadlevel-0-runbook.md (pass A -> Taxonomy, pass B -> Classification)." & vbCrLf & vbCrLf & _
           "File: " & outPath, vbInformation, "Threads for LLM"
End Sub

Private Sub BuildThreadsSheet(ByVal ws As Object, ByRef tData() As Variant, ByVal tCount As Long)
    Dim hdrs As Variant
    hdrs = Array("ThreadId", "ConversationTopic", "FirstInboundTime", "LastMessageTime", _
                 "LastMessageDirection", "MsgCount", "InboundCount", "OutboundCount", _
                 "OffHoursTrigger", "AnyInboundOffHours", "AutomatedSender", "DeadEndHeuristic", _
                 "HasAPS", "HasADM", "HasDEV", "RecurrenceGroupId", "RecurrenceCount", _
                 "NormalizedSubject", "SampleForTaxonomy", "CondensedText")
    Dim hi As Long
    For hi = 0 To UBound(hdrs)
        ws.Cells(1, hi + 1).Value = hdrs(hi)
    Next hi
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 20)).Font.Bold = True
    ws.Range(ws.Cells(2, 1), ws.Cells(tCount + 1, 20)).Value = tData
    ws.Range(ws.Cells(1, 1), ws.Cells(tCount + 1, 20)).Sort _
        Key1:=ws.Range("C2"), Order1:=1, Header:=1
    ws.Range(ws.Cells(1, 1), ws.Cells(tCount + 1, 20)).AutoFilter
    ws.Columns("A:S").AutoFit
    ws.Columns("T").ColumnWidth = 100
End Sub

Private Sub BuildMessagesSheet(ByVal ws As Object, ByVal msgCount As Long, ByVal maxRows As Long)
    ws.Name = "Messages"
    Dim hdrs As Variant
    hdrs = Array("ThreadId", "MessageId", "IsDuplicateMsgId", "Direction", "When", "IsOffHours", _
                 "From", "FromTeam", "AutomatedSender", "HasAPS", "HasADM", "HasDEV", _
                 "Subject", "Body", "FolderPath", "EntryID")
    Dim hi As Long
    For hi = 0 To UBound(hdrs)
        ws.Cells(1, hi + 1).Value = hdrs(hi)
    Next hi
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 16)).Font.Bold = True

    Dim d() As Variant
    ReDim d(1 To msgCount, 1 To 16)
    Dim i As Long
    For i = 1 To msgCount
        d(i, 1) = mConv(i): d(i, 2) = mMsgId(i): d(i, 3) = mIsDup(i): d(i, 4) = mDir(i)
        d(i, 5) = Format(mWhen(i), "yyyy-mm-dd hh:nn:ss"): d(i, 6) = mOff(i)
        d(i, 7) = mFrom(i): d(i, 8) = mFromTeam(i): d(i, 9) = mAuto(i)
        d(i, 10) = mHasAPS(i): d(i, 11) = mHasADM(i): d(i, 12) = mHasDEV(i)
        d(i, 13) = mSubject(i): d(i, 14) = mBody(i): d(i, 15) = mFolder(i): d(i, 16) = mEntry(i)
    Next i
    ws.Range(ws.Cells(2, 1), ws.Cells(msgCount + 1, 16)).Value = d
    ws.Range(ws.Cells(1, 1), ws.Cells(msgCount + 1, 16)).Sort _
        Key1:=ws.Range("A2"), Order1:=1, Key2:=ws.Range("E2"), Order2:=1, Header:=1
    ws.Range(ws.Cells(1, 1), ws.Cells(msgCount + 1, 16)).AutoFilter
    ws.Columns("A:M").AutoFit
    ws.Columns("N").ColumnWidth = 80
End Sub

Private Sub BuildStatsSheet(ByVal ws As Object, ByVal outPath As String, ByVal windowDesc As String, _
                            ByVal folderDesc As String, ByVal total As Long, ByVal msgCount As Long, _
                            ByVal tCount As Long, ByVal offHoursCount As Long, ByVal dupCount As Long, _
                            ByVal inboundCount As Long, ByVal outboundCount As Long, ByVal recurGroups As Long)
    ws.Name = "Stats"
    ws.Cells(1, 1).Value = "Threads for LLM (v2) - Run Statistics"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 12
    Dim r As Long
    r = 3
    AddStatRow ws, r, "Run timestamp", Format(Now, "yyyy-mm-dd hh:nn:ss")
    AddStatRow ws, r, "Mailbox", IIf(Len(SHARED_MAILBOX) = 0, "<your own>", SHARED_MAILBOX)
    AddStatRow ws, r, "Folders scanned", folderDesc & IIf(SCAN_SENT_ALSO, " (Inbox + Sent)", "")
    AddStatRow ws, r, "Scan window", windowDesc
    AddStatRow ws, r, "Off-hours rule", "original Paris-gated (current coverage gap)"
    AddStatRow ws, r, "Direction basis", "sender (SUPPORT_ADDRESSES_CSV); Sent folder " & IIf(SCAN_SENT_ALSO, "scanned", "not scanned")
    AddStatRow ws, r, "Output file", outPath
    r = r + 1
    SectionHeader ws, r, "Volume"
    AddStatRow ws, r, "Messages exported", msgCount
    AddStatRow ws, r, "Threads (conversations)", tCount
    AddStatRow ws, r, "Recurrence groups", recurGroups
    AddStatRow ws, r, "Inbound / Outbound messages", inboundCount & " / " & outboundCount
    AddStatRow ws, r, "Off-hours messages", offHoursCount
    AddStatRow ws, r, "Exact duplicates flagged (Message-ID)", dupCount
    If outboundCount = 0 Then
        r = r + 1
        ws.Cells(r, 1).Value = "WARNING: 0 Outbound messages. Set SUPPORT_ADDRESSES_CSV (and/or check that " & _
            "replies are CC'd to this inbox) - without support replies the resolvability signal is lost."
        ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 1).Font.Color = RGB(200, 60, 60)
    End If
    If Len(APS_ADDRESSES_CSV) = 0 And Len(ADM_ADDRESSES_CSV) = 0 And Len(DEV_ADDRESSES_CSV) = 0 Then
        r = r + 1
        ws.Cells(r, 1).Value = "(team rosters empty - LLM infers teams from content only)"
        ws.Cells(r, 1).Font.Italic = True
        ws.Cells(r, 1).Font.Color = RGB(150, 150, 150)
    End If
    ws.Columns("A:B").AutoFit
End Sub

Private Sub BuildTaxonomySheet(ByVal ws As Object)
    ws.Name = "Taxonomy"
    Dim hdrs As Variant
    hdrs = Array("category_id", "name", "description", "default_team", "default_aps_resolvable", "default_time_sensitive")
    Dim hi As Long
    For hi = 0 To UBound(hdrs)
        ws.Cells(1, hi + 1).Value = hdrs(hi)
    Next hi
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 6)).Font.Bold = True
    ws.Cells(2, 1).Value = "<paste the table from prompts/threadlevel-1-induce-taxonomy.md here>"
    ws.Cells(2, 1).Font.Italic = True
    ws.Cells(2, 1).Font.Color = RGB(150, 150, 150)
    ws.Columns("A:F").AutoFit
End Sub

Private Sub BuildClassificationSheet(ByVal ws As Object)
    ws.Name = "Classification"
    Dim hdrs As Variant
    hdrs = Array("ThreadId", "category_id", "message_type", "is_genuine_issue", "dead_end", _
                 "offhours_trigger", "resolver_team", "aps_resolvable_alone", "time_sensitive", _
                 "confidence", "needs_review", "rationale")
    Dim hi As Long
    For hi = 0 To UBound(hdrs)
        ws.Cells(1, hi + 1).Value = hdrs(hi)
    Next hi
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 12)).Font.Bold = True
    ws.Cells(2, 1).Value = "<paste pass-B rows from prompts/threadlevel-2-classify-threads.md here (one per thread)>"
    ws.Cells(2, 1).Font.Italic = True
    ws.Cells(2, 1).Font.Color = RGB(150, 150, 150)
    ws.Columns("A:L").AutoFit
End Sub

Private Sub BuildSummarySheet(ByVal ws As Object)
    ws.Name = "Summary"
    ws.Cells(1, 1).Value = "Follow-the-sun Summary (counts computed from the Classification sheet)"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 12

    ' Classification cols: A ThreadId, C message_type, D is_genuine_issue, E dead_end,
    ' F offhours_trigger, H aps_resolvable_alone, I time_sensitive, K needs_review
    Dim genuine As String
    genuine = "Classification!D:D,""yes"",Classification!F:F,""yes"",Classification!E:E,""no"""

    SetLabelFormula ws, 3, "Classified threads", "=COUNTA(Classification!A:A)-1"
    SetLabelFormula ws, 4, "Genuine off-hours issues", "=COUNTIFS(" & genuine & ")"
    SetLabelFormula ws, 5, "  ...APS-resolvable alone", "=COUNTIFS(" & genuine & ",Classification!H:H,""yes"")"
    SetLabelFormula ws, 6, "  ...APS-alone AND time-sensitive", "=COUNTIFS(" & genuine & ",Classification!H:H,""yes"",Classification!I:I,""yes"")"
    SetLabelFormula ws, 7, "APS-alone share of genuine off-hours", "=IFERROR(B5/B4,0)"
    SetLabelFormula ws, 8, "Follow-the-sun value share", "=IFERROR(B6/B4,0)"
    SetLabelFormula ws, 9, "Needs review", "=COUNTIF(Classification!K:K,""yes"")"
    ws.Cells(7, 2).NumberFormat = "0.0%"
    ws.Cells(8, 2).NumberFormat = "0.0%"

    ws.Cells(11, 1).Value = "VERDICT: large 'follow-the-sun value share' => 24h APS coverage is MEANINGFUL; " & _
                            "small => MARGINAL (mostly dead-ends, noise, or needed DEV/ADM)."
    ws.Cells(12, 1).Value = "For per-category / per-team / recurrence breakdowns, insert a PivotTable on the " & _
                            "Threads + Classification sheets (group by category_id, resolver_team, RecurrenceCount)."
    ws.Cells(11, 1).Font.Italic = True
    ws.Cells(12, 1).Font.Italic = True
    ws.Columns("A:A").ColumnWidth = 42
    ws.Columns("B:B").AutoFit
End Sub

Private Sub BuildGoldSetSheet(ByVal ws As Object)
    ws.Name = "GoldSet"
    ws.Cells(1, 1).Value = "GoldSet calibration - paste ~30 ThreadIds (col A) and your own labels (cols B,C). Model cols auto-fill."
    ws.Cells(1, 1).Font.Bold = True
    SetLabelFormula ws, 2, "Agreement message_type", "=IFERROR(AVERAGE(F5:F34),"""")"
    SetLabelFormula ws, 3, "Agreement aps_resolvable", "=IFERROR(AVERAGE(G5:G34),"""")"
    ws.Cells(2, 2).NumberFormat = "0.0%"
    ws.Cells(3, 2).NumberFormat = "0.0%"

    Dim hdrs As Variant
    hdrs = Array("ThreadId", "human_message_type", "human_aps_resolvable", _
                 "model_message_type", "model_aps_resolvable", "agree_type", "agree_aps")
    Dim hi As Long
    For hi = 0 To UBound(hdrs)
        ws.Cells(4, hi + 1).Value = hdrs(hi)
    Next hi
    ws.Range(ws.Cells(4, 1), ws.Cells(4, 7)).Font.Bold = True

    Dim rr As Long
    For rr = 5 To 34
        ws.Cells(rr, 4).Formula = "=IFERROR(VLOOKUP(A" & rr & ",Classification!$A:$L,3,FALSE),"""")"
        ws.Cells(rr, 5).Formula = "=IFERROR(VLOOKUP(A" & rr & ",Classification!$A:$L,8,FALSE),"""")"
        ws.Cells(rr, 6).Formula = "=IF(AND($B" & rr & "<>"""",$D" & rr & "<>""""),--(LOWER($B" & rr & ")=LOWER($D" & rr & ")),"""")"
        ws.Cells(rr, 7).Formula = "=IF(AND($C" & rr & "<>"""",$E" & rr & "<>""""),--(LOWER($C" & rr & ")=LOWER($E" & rr & ")),"""")"
    Next rr
    ws.Columns("A:G").AutoFit
End Sub

Private Sub SetLabelFormula(ByVal ws As Object, ByVal r As Long, ByVal label As String, ByVal formula As String)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Formula = formula
End Sub


' ----------------------- Thread-feature helpers -----------------------

' A message is "non-actionable" for the dead-end test when it is an outbound
' reply, a short thanks/ack, or a short inbound that arrives after the last
' support reply (i.e. trailing pleasantry on an already-handled thread).
Private Function IsNonActionable(ByVal idx As Long, ByVal lastOutTime As Date) As Boolean
    If mDir(idx) = "Outbound" Then IsNonActionable = True : Exit Function
    If ThanksLike(mBody(idx)) And Len(mBody(idx)) <= DEADEND_MAX_CHARS Then IsNonActionable = True : Exit Function
    If lastOutTime > 0 And mWhen(idx) > lastOutTime And Len(mBody(idx)) <= DEADEND_MAX_CHARS Then
        IsNonActionable = True
    End If
End Function

Private Function ThanksLike(ByVal body As String) As Boolean
    Dim s As String, toks As Variant, i As Long
    s = LCase(body)
    toks = Array("thank", "thx", "thanks", "dzięk", "dziek", "merci", "cheers", _
                 "appreciate", "received", "got it", "noted", "ok ", "okay")
    For i = LBound(toks) To UBound(toks)
        If InStr(s, toks(i)) > 0 Then ThanksLike = True : Exit Function
    Next i
End Function

' Strip reply/forward prefixes (incl. PL/DE/FR) and bracketed ticket ids; lower/trim.
Private Function NormalizeSubject(ByVal s As String) As String
    Dim t As String
    t = Trim(s)
    Dim changed As Boolean
    Do
        changed = False
        Dim pfx As Variant, p As Variant
        pfx = Array("re:", "fw:", "fwd:", "aw:", "wg:", "odp:", "re :", "fwd :")
        For Each p In pfx
            If LCase(Left(LTrim(t), Len(p))) = p Then
                t = Mid(LTrim(t), Len(p) + 1)
                changed = True
            End If
        Next p
    Loop While changed
    ' drop a leading [TICKET-123] style tag
    Do While Left(LTrim(t), 1) = "["
        Dim cb As Long
        cb = InStr(t, "]")
        If cb = 0 Then Exit Do
        t = Mid(t, cb + 1)
    Loop
    NormalizeSubject = Trim(LCase(t))
End Function


' ----------------------- Timing -----------------------

Private Function ComputeOffHours(ByVal t As Date) As Boolean
    Dim isWknd As Boolean, hourLocal As Integer, minuteOfDay As Long, outsideOffice As Boolean
    Dim hPL As Boolean, hIN As Boolean, indiaEndMin As Long
    isWknd = IsWeekendDay(t)
    hourLocal = Hour(t)
    minuteOfDay = hourLocal * 60 + Minute(t)
    outsideOffice = (minuteOfDay < WORK_START_HOUR * 60 Or minuteOfDay >= WORK_END_HOUR * 60)
    hPL = IsHolidayPL(t)
    hIN = IsHolidayIN(t)
    If IsSummerTime(t) Then indiaEndMin = INDIA_END_SUMMER_MIN Else indiaEndMin = INDIA_END_WINTER_MIN
    ComputeOffHours = (isWknd And TREAT_SATURDAY_AS_OFFHOURS And Weekday(t, vbMonday) = 6) Or _
                      (isWknd And TREAT_SUNDAY_AS_OFFHOURS And Weekday(t, vbMonday) = 7) Or _
                      outsideOffice Or _
                      (hPL And (hIN Or minuteOfDay >= indiaEndMin))
End Function


' ============================================================================
'  Helpers copied from extract-emails-for-llm.bas (self-contained module).
'  Keep holiday CSVs + timing helpers in sync across all macros.
' ============================================================================

Private Function ResolveWindow(ByRef fromDt As Date, ByRef toDt As Date, ByRef windowDesc As String) As Boolean
    If Len(SCAN_FROM_DATE) > 0 Then
        On Error Resume Next
        fromDt = CDate(SCAN_FROM_DATE)
        If Err.Number <> 0 Then
            MsgBox "SCAN_FROM_DATE could not be parsed: """ & SCAN_FROM_DATE & """ (use YYYY-MM-DD).", vbExclamation, "Threads for LLM"
            On Error GoTo 0: Exit Function
        End If
        On Error GoTo 0
    ElseIf SCAN_LAST_N_DAYS > 0 Then
        fromDt = DateAdd("d", -SCAN_LAST_N_DAYS, Now)
    End If
    If Len(SCAN_TO_DATE) > 0 Then
        On Error Resume Next
        toDt = CDate(SCAN_TO_DATE) + TimeSerial(23, 59, 59)
        If Err.Number <> 0 Then
            MsgBox "SCAN_TO_DATE could not be parsed: """ & SCAN_TO_DATE & """ (use YYYY-MM-DD).", vbExclamation, "Threads for LLM"
            On Error GoTo 0: Exit Function
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

Private Function EarliestMarker(ByVal s As String) As Long
    Dim markers As Variant, i As Long, p As Long, best As Long
    markers = Array("-----Original Message-----", "________________________________", _
                    vbCrLf & "From: ", vbLf & "From: ", "-----Wiadomo", "-------- Forwarded Message --------")
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
        If Left(LTrim(ln), 1) <> ">" Then out = out & ln & vbLf
    Next i
    RemoveQuotedLines = out
End Function

Private Function CollapseBlankLines(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    Do While InStr(s, vbLf & vbLf & vbLf) > 0
        s = Replace(s, vbLf & vbLf & vbLf, vbLf & vbLf)
    Loop
    CollapseBlankLines = s
End Function

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

Private Function TeamOf(ByVal fromLower As String) As String
    If ContainsAnyToken(fromLower, APS_ADDRESSES_CSV) Then TeamOf = "APS" : Exit Function
    If ContainsAnyToken(fromLower, ADM_ADDRESSES_CSV) Then TeamOf = "ADM" : Exit Function
    If ContainsAnyToken(fromLower, DEV_ADDRESSES_CSV) Then TeamOf = "DEV" : Exit Function
    TeamOf = "Other"
End Function

' True when the sender is OUR support side => this message is an Outbound reply
' (even a CC'd copy sitting in the Inbox). Matches SUPPORT_ADDRESSES_CSV plus the
' shared mailbox SMTP address when SHARED_MAILBOX is an address.
Private Function IsSupportSender(ByVal fromLower As String) As Boolean
    If ContainsAnyToken(fromLower, SUPPORT_ADDRESSES_CSV) Then IsSupportSender = True : Exit Function
    If InStr(SHARED_MAILBOX, "@") > 0 Then
        If InStr(fromLower, LCase(SHARED_MAILBOX)) > 0 Then IsSupportSender = True
    End If
End Function

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
        MsgBox "Cannot create output directory:" & vbCrLf & outDir, vbCritical, "Threads for LLM"
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
