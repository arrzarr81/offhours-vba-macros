Attribute VB_Name = "OffHoursEmailsIndiaInHours"
' ============================================================================
'  Off-Hours Email Finder for Outlook
'  -- VARIANT: India team's shift counts as IN business hours --
'  Project: p5-outlook-offhours-emails (APS-llm-tools)
'
'  HOW THIS DIFFERS FROM find-offhours-emails.bas
'  -----------------------------------------------------------------
'  The original script defines business hours as PARIS 09:00-18:00 only.
'  India ever only EXTENDS coverage in the early afternoon of a Polish
'  holiday (the "India coverage end" cutoff). On a normal weekday an email
'  at 07:00 Paris -- when the India team (IST, UTC+5:30) is already at their
'  desks -- is still flagged off-hours, because it is before Paris 09:00.
'
'  THIS variant treats the India team as a full first-class support shift.
'  Business hours = (Paris team working) OR (India team working), every day.
'  An email is OFF-HOURS only when BOTH teams are off. So 07:00 Paris on a
'  normal weekday (~11:30 IST) is now IN-HOURS, because India is staffed.
'
'  Both teams modelled symmetrically:
'    - Paris team: Mon-Fri, 09:00-18:00 Paris local, OFF on Polish holidays
'                  (Poland is the gating calendar, exactly as the original).
'    - India team: Mon-Fri, INDIA_SHIFT_*_IST_MIN in IST, OFF on Mumbai
'                  holidays. The email's Paris time is converted to IST
'                  (winter +4:30, summer +3:30 -- the existing DST logic)
'                  and tested against the India shift window.
'
'  IsOffHours = NOT ( ParisTeamWorking OR IndiaTeamWorking )
'
'  This naturally subsumes the original's "India coverage end" rule: on a
'  Polish holiday, the morning/early afternoon is covered by India until
'  18:00 IST (= 13:30 Paris winter / 14:30 summer), after which both teams
'  are off and it becomes off-hours -- same boundary, derived instead of
'  hard-coded.
'
'  France holidays (IsHolidayFR) are tracked as a column for visibility only;
'  they do NOT drive the decision (no French-team gating in this model).
'
'  Everything else (read-only safety, shared-mailbox handling, subfolder
'  traversal, scan window, Users-tab IsAppUser join) is identical to the
'  original. Output filename is prefixed "offhours-emails-india" so the two
'  scripts' XLSX outputs never collide.
'
'  READ-ONLY. NO WRITE CAPABILITY. (Same structural guarantee as the
'  original: no .Save / Categories / UnRead / Move / Delete / Send anywhere.)
'
'  How to use:
'    1. Outlook -> Alt+F11 -> File -> Import File... -> pick this .bas
'    2. Set INDIA_SHIFT_START_IST_MIN / INDIA_SHIFT_END_IST_MIN to your
'       India team's ACTUAL shift, plus the usual CONFIG constants below.
'    3. Run macro: FindOffHoursEmailsIndiaInHours (F5)
' ============================================================================

Option Explicit

' -------------------- CONFIG (edit these) --------------------
' Paris team working window (local time). In-hours when
' WORK_START_HOUR <= Hour(ReceivedTime) < WORK_END_HOUR, Mon-Fri.
Private Const WORK_START_HOUR As Integer = 9            ' inclusive, Paris local (09:00)
Private Const WORK_END_HOUR As Integer = 18            ' exclusive, Paris local (18:00)

' India (Mumbai) team shift, expressed in IST (Asia/Kolkata, fixed UTC+5:30,
' no DST). Minutes since IST midnight. EDIT to match your India team's real
' rostered hours. Defaults: 09:00-18:00 IST.
'
' The shift is FIXED in IST and therefore floats in Paris time across DST:
'   09:00 IST start == 05:30 CEST (summer) / 04:30 CET (winter)
'   18:00 IST end   == 14:30 CEST (summer) / 13:30 CET (winter)
'                       (end matches the original script's coverage cutoff)
' Confirmed: India currently starts 05:30 CEST in summer == 09:00 IST. Because
' IST has no DST, that same 09:00 IST appears as 04:30 CET in winter. If your
' India team instead shifts their IST hours to hold 05:30 Paris year-round,
' this fixed-IST model is NOT what you want - say so and it can be reworked.
Private Const INDIA_SHIFT_START_IST_MIN As Long = 9 * 60   ' 09:00 IST = 540  (05:30 CEST / 04:30 CET)
Private Const INDIA_SHIFT_END_IST_MIN As Long = 18 * 60    ' 18:00 IST = 1080 (14:30 CEST / 13:30 CET)
Private Const INDIA_WORKS_SATURDAY As Boolean = False      ' set True if the India team rosters Saturdays

' IST is ahead of Paris by 4h30 in winter (CET, UTC+1) and 3h30 in summer
' (CEST, UTC+2). Used to convert ReceivedTime (Paris local) -> IST.
Private Const IST_OFFSET_WINTER_MIN As Long = 4 * 60 + 30  ' +4:30 = 270
Private Const IST_OFFSET_SUMMER_MIN As Long = 3 * 60 + 30  ' +3:30 = 210

' Public holidays (YYYY-MM-DD, comma-separated). KEEP IN SYNC with
' find-offhours-emails.bas and add-timing-columns.bas.
'   - HOLIDAYS_PL gates the Paris/Poland team (Poland off => Paris team off).
'   - HOLIDAYS_IN gates the India team (Mumbai off => India team off).
'   - HOLIDAYS_FR is tracked for the IsHolidayFR column only; no gating.
Private Const HOLIDAYS_FR_CSV As String = _
    "2025-01-01," & _
    "2025-04-21," & _
    "2025-05-01,2025-05-08,2025-05-29," & _
    "2025-06-09," & _
    "2025-07-14," & _
    "2025-08-15," & _
    "2025-11-01,2025-11-11," & _
    "2025-12-25," & _
    "2026-01-01," & _
    "2026-04-06," & _
    "2026-05-01,2026-05-08,2026-05-14,2026-05-25," & _
    "2026-07-14," & _
    "2026-08-15," & _
    "2026-11-01,2026-11-11," & _
    "2026-12-25"

Private Const HOLIDAYS_PL_CSV As String = _
    "2025-01-01,2025-01-06," & _
    "2025-04-20,2025-04-21," & _
    "2025-05-01,2025-05-03," & _
    "2025-06-08,2025-06-19," & _
    "2025-08-15," & _
    "2025-11-01,2025-11-11," & _
    "2025-12-24,2025-12-25,2025-12-26," & _
    "2026-01-01,2026-01-06," & _
    "2026-04-05,2026-04-06," & _
    "2026-05-01,2026-05-03,2026-05-24," & _
    "2026-06-04," & _
    "2026-08-15," & _
    "2026-11-01,2026-11-11," & _
    "2026-12-24,2026-12-25,2026-12-26"

Private Const HOLIDAYS_IN_MUMBAI_CSV As String = _
    "2025-01-26,2025-02-19," & _
    "2025-03-14,2025-03-31," & _
    "2025-04-10,2025-04-14,2025-04-18," & _
    "2025-05-01,2025-05-12," & _
    "2025-06-07," & _
    "2025-07-06," & _
    "2025-08-15,2025-08-27," & _
    "2025-10-02,2025-10-21,2025-10-22," & _
    "2025-11-05," & _
    "2025-12-25," & _
    "2026-01-26," & _
    "2026-03-04,2026-03-19," & _
    "2026-04-03,2026-04-14," & _
    "2026-05-01," & _
    "2026-08-15," & _
    "2026-09-14," & _
    "2026-10-02,2026-10-20," & _
    "2026-11-08,2026-11-09," & _
    "2026-12-25"

' Mailbox to scan.
'   ""                       = your own default mailbox
'   "Project Inbox Name"     = a shared mailbox by display name
'   "project@company.com"    = a shared mailbox by SMTP address
Private Const SHARED_MAILBOX As String = ""

' Folder name within the resolved mailbox.
Private Const FOLDER_NAME As String = "Inbox"

' Recursive subfolder traversal (False = scan only FOLDER_NAME).
Private Const INCLUDE_SUBFOLDERS As Boolean = False

' --- Scan window (pick ONE pattern; see original for details) ---
Private Const SCAN_FROM_DATE As String = ""
Private Const SCAN_TO_DATE As String = ""
Private Const SCAN_LAST_N_DAYS As Long = 30

' Output location. "" = <Documents>\OffHoursEmails\
Private Const OUTPUT_DIR As String = ""
Private Const ARCHIVE_FILES As Boolean = True
Private Const FILENAME_PREFIX As String = "offhours-emails-india"
' -------------------------------------------------------------


Public Sub FindOffHoursEmailsIndiaInHours()
    Dim ns As Outlook.NameSpace
    Dim folder As Outlook.folder
    Dim itm As Object
    Dim total As Long, offHoursCount As Long
    Dim countHolidayFR As Long, countHolidayPL As Long, countHolidayIN As Long
    Dim countParisHours As Long, countIndiaHours As Long, countIndiaRescued As Long

    Set ns = Application.GetNamespace("MAPI")
    Set folder = ResolveFolder(ns)
    If folder Is Nothing Then
        MsgBox "Folder not found. Mailbox=""" & SHARED_MAILBOX & """, Folder=""" & FOLDER_NAME & """", vbExclamation
        Exit Sub
    End If

    Dim outPath As String
    outPath = ResolveOutputPath()
    If Len(outPath) = 0 Then Exit Sub

    ' Resolve the scan window: explicit date range first, fallback to N days back.
    Dim fromDt As Date, toDt As Date
    Dim windowDesc As String

    If Len(SCAN_FROM_DATE) > 0 Then
        On Error Resume Next
        fromDt = CDate(SCAN_FROM_DATE)
        If Err.Number <> 0 Then
            MsgBox "SCAN_FROM_DATE could not be parsed: """ & SCAN_FROM_DATE & """" & vbCrLf & _
                   "Use YYYY-MM-DD (e.g. 2025-04-01).", vbExclamation, "Off-Hours Email Finder (India in-hours)"
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    ElseIf SCAN_LAST_N_DAYS > 0 Then
        fromDt = DateAdd("d", -SCAN_LAST_N_DAYS, Now)
    End If

    If Len(SCAN_TO_DATE) > 0 Then
        On Error Resume Next
        toDt = CDate(SCAN_TO_DATE) + TimeSerial(23, 59, 59)
        If Err.Number <> 0 Then
            MsgBox "SCAN_TO_DATE could not be parsed: """ & SCAN_TO_DATE & """" & vbCrLf & _
                   "Use YYYY-MM-DD (e.g. 2025-04-30).", vbExclamation, "Off-Hours Email Finder (India in-hours)"
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    End If

    ' Build Outlook restriction filter (server/index-side).
    Dim restriction As String
    If fromDt > 0 Then
        restriction = "[ReceivedTime] >= '" & Format(fromDt, "ddddd h:nn AMPM") & "'"
    End If
    If toDt > 0 Then
        If Len(restriction) > 0 Then restriction = restriction & " AND "
        restriction = restriction & "[ReceivedTime] <= '" & Format(toDt, "ddddd h:nn AMPM") & "'"
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

    ' Build the list of folders to scan (root, or root + descendants).
    Dim folderList As Collection, pathList As Collection
    Set folderList = New Collection
    Set pathList = New Collection
    CollectFolders folder, "", folderList, pathList, INCLUDE_SUBFOLDERS

    Dim maxRows As Long
    Dim fIdx As Long
    maxRows = 0
    For fIdx = 1 To folderList.Count
        Dim fTmp As Outlook.folder
        Set fTmp = folderList(fIdx)
        maxRows = maxRows + fTmp.items.Count
    Next fIdx
    If maxRows < 1 Then maxRows = 1

    ' Pre-allocate two buffers. Final sheet layout:
    '   A:K = leftData  (timing block, 11 cols)
    '   L   = IsAppUser formula
    '   M:P = rightData (From, Subject, EntryID, FolderPath)
    Dim leftData() As Variant   ' 11 cols
    Dim rightData() As Variant  ' 4 cols
    ReDim leftData(1 To maxRows, 1 To 11)
    ReDim rightData(1 To maxRows, 1 To 4)
    Dim row As Long
    row = 0

    For fIdx = 1 To folderList.Count
        Dim curFolder As Outlook.folder
        Set curFolder = folderList(fIdx)
        Dim curPath As String
        curPath = pathList(fIdx)

        Dim folderScan As Outlook.items
        Set folderScan = curFolder.items
        If Len(restriction) > 0 Then
            Set folderScan = folderScan.Restrict(restriction)
        End If
        folderScan.Sort "[ReceivedTime]", True   ' newest first

        For Each itm In folderScan
            If TypeOf itm Is Outlook.MailItem Then
                Dim m As Outlook.MailItem
                Set m = itm
                total = total + 1

                Dim t As Date
                t = m.ReceivedTime

                ' --- Holiday flags (visibility columns) ---
                Dim hFR As Boolean, hPL As Boolean, hIN As Boolean
                hFR = IsHolidayFR(t)
                hPL = IsHolidayPL(t)

                ' --- Paris team working? Mon-Fri, in window, not a PL holiday ---
                Dim parisDow As Integer, parisMin As Long
                Dim isWknd As Boolean, parisWorking As Boolean
                parisDow = Weekday(t, vbMonday)          ' 1=Mon .. 7=Sun
                parisMin = Hour(t) * 60 + Minute(t)
                isWknd = (parisDow = 6 Or parisDow = 7)
                parisWorking = (parisDow <= 5) And (Not hPL) And _
                               (parisMin >= WORK_START_HOUR * 60 And parisMin < WORK_END_HOUR * 60)

                ' --- India team working? Convert Paris time -> IST, test shift ---
                Dim istTime As Date, offsetMin As Long
                If IsSummerTime(t) Then
                    offsetMin = IST_OFFSET_SUMMER_MIN
                Else
                    offsetMin = IST_OFFSET_WINTER_MIN
                End If
                istTime = t + offsetMin / 1440#          ' add minutes as fraction of a day

                ' Mumbai holiday is evaluated on the IST calendar date (handles
                ' the late-evening-Paris roll into the next IST day).
                hIN = IsHolidayIN(istTime)

                Dim istDow As Integer, istMin As Long, indiaWorking As Boolean
                Dim indiaIsWorkday As Boolean
                istDow = Weekday(istTime, vbMonday)
                istMin = Hour(istTime) * 60 + Minute(istTime)
                indiaIsWorkday = (istDow <= 5) Or (istDow = 6 And INDIA_WORKS_SATURDAY)
                indiaWorking = indiaIsWorkday And (Not hIN) And _
                               (istMin >= INDIA_SHIFT_START_IST_MIN And istMin < INDIA_SHIFT_END_IST_MIN)

                ' --- Off-hours = neither team on the clock ---
                Dim isOff As Boolean
                isOff = Not (parisWorking Or indiaWorking)

                ' Tallies
                If isOff Then offHoursCount = offHoursCount + 1
                If hFR Then countHolidayFR = countHolidayFR + 1
                If hPL Then countHolidayPL = countHolidayPL + 1
                If hIN Then countHolidayIN = countHolidayIN + 1
                If parisWorking Then countParisHours = countParisHours + 1
                If indiaWorking Then countIndiaHours = countIndiaHours + 1
                ' "Rescued" = would be off-hours on Paris-only hours, but India
                ' coverage makes it in-hours -- the delta vs. the original script.
                If indiaWorking And Not parisWorking Then countIndiaRescued = countIndiaRescued + 1

                row = row + 1
                leftData(row, 1) = Format(t, "yyyy-mm-dd hh:nn:ss")       ' ReceivedTime (Paris)
                leftData(row, 2) = Hour(t)                                ' HourLocal (Paris)
                leftData(row, 3) = Hour(istTime)                          ' HourIST (India)
                leftData(row, 4) = WeekdayName(parisDow, True)            ' Weekday (Paris)
                leftData(row, 5) = isWknd                                 ' IsWeekend
                leftData(row, 6) = hFR                                    ' IsHolidayFR (info only)
                leftData(row, 7) = hPL                                    ' IsHolidayPL (gates Paris)
                leftData(row, 8) = hIN                                    ' IsHolidayIN (gates India)
                leftData(row, 9) = parisWorking                           ' IsParisHours
                leftData(row, 10) = indiaWorking                          ' IsIndiaHours
                leftData(row, 11) = isOff                                 ' IsOffHours
                rightData(row, 1) = SafeSenderAddress(m)                  ' From
                rightData(row, 2) = m.Subject                            ' Subject
                rightData(row, 3) = m.EntryID                            ' EntryID
                rightData(row, 4) = curPath                              ' FolderPath
            End If
        Next itm
    Next fIdx

    If row = 0 Then
        MsgBox "No mails in the scan window. No file written.", vbInformation, "Off-Hours Email Finder (India in-hours)"
        Exit Sub
    End If

    ' --- Write XLSX via Excel automation ---
    Dim xl As Object
    On Error Resume Next
    Set xl = CreateObject("Excel.Application")
    On Error GoTo 0
    If xl Is Nothing Then
        MsgBox "Could not start Excel. The macro requires Excel to write the XLSX. " & _
               "Confirm Excel is installed and try again.", vbCritical, "Off-Hours Email Finder (India in-hours)"
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

    ' --- Sheet 1: "Emails" ---
    Dim ws As Object
    Set ws = wb.Worksheets(1)
    ws.Name = "Emails"

    Dim hdrs As Variant
    hdrs = Array("ReceivedTime", "HourLocal", "HourIST", "Weekday", "IsWeekend", _
                 "IsHolidayFR", "IsHolidayPL", "IsHolidayIN", _
                 "IsParisHours", "IsIndiaHours", "IsOffHours", _
                 "IsAppUser", "From", "Subject", "EntryID", "FolderPath")
    Dim hi As Long
    For hi = 0 To UBound(hdrs)
        ws.Cells(1, hi + 1).Value = hdrs(hi)
    Next hi
    ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(hdrs) + 1)).Font.Bold = True

    ' Sheet layout:
    '   A:K = leftData (timing, 11 cols)
    '   L   = IsAppUser formula
    '   M:P = rightData (From, Subject, EntryID, FolderPath)
    ws.Range(ws.Cells(2, 1), ws.Cells(maxRows + 1, 11)).Value = leftData
    ws.Range(ws.Cells(2, 13), ws.Cells(maxRows + 1, 16)).Value = rightData
    If row < maxRows Then
        ws.Range(ws.Cells(row + 2, 1), ws.Cells(maxRows + 1, 11)).ClearContents
        ws.Range(ws.Cells(row + 2, 13), ws.Cells(maxRows + 1, 16)).ClearContents
    End If

    ' IsAppUser formula in col L for populated rows only. From is in col M.
    ws.Range(ws.Cells(2, 12), ws.Cells(row + 1, 12)).Formula = _
        "=IF(M2="""",FALSE,COUNTIF(Users!A:A,M2)>0)"

    ' AutoFilter on the full data range (A:P)
    ws.Range(ws.Cells(1, 1), ws.Cells(row + 1, 16)).AutoFilter

    ws.Columns("A:P").AutoFit

    ' --- Sheet 2: "Users" ---
    Dim wsUsers As Object
    Set wsUsers = wb.Worksheets.Add(After:=ws)
    wsUsers.Name = "Users"
    wsUsers.Cells(1, 1).Value = "Email"
    wsUsers.Cells(1, 2).Value = "Notes (optional)"
    wsUsers.Range(wsUsers.Cells(1, 1), wsUsers.Cells(1, 2)).Font.Bold = True
    wsUsers.Cells(2, 1).Value = "Paste your application user emails here, one per row in column A."
    wsUsers.Cells(2, 1).Font.Italic = True
    wsUsers.Cells(2, 1).Font.Color = RGB(120, 120, 120)
    wsUsers.Cells(3, 1).Value = "(this hint row does not match real email addresses, so it is ignored)"
    wsUsers.Cells(3, 1).Font.Italic = True
    wsUsers.Cells(3, 1).Font.Color = RGB(160, 160, 160)
    wsUsers.Columns("A:B").AutoFit

    ' Human-readable folder description (also reused in the summary dialog below)
    Dim folderDesc As String
    If INCLUDE_SUBFOLDERS And folderList.Count > 1 Then
        folderDesc = FOLDER_NAME & " + " & (folderList.Count - 1) & " subfolders"
    Else
        folderDesc = FOLDER_NAME
    End If

    ' --- Sheet 3: "Stats" - persisted run statistics. One Stats tab per run,
    '     since every run writes a fresh timestamped workbook (ARCHIVE_FILES). ---
    Dim wsStats As Object
    Set wsStats = wb.Worksheets.Add(After:=ws)
    wsStats.Name = "Stats"
    wsStats.Cells(1, 1).Value = "Off-Hours Email Finder - Run Statistics (India counts as in-hours)"
    wsStats.Cells(1, 1).Font.Bold = True
    wsStats.Cells(1, 1).Font.Size = 12

    Dim r As Long
    r = 3
    AddStatRow wsStats, r, "Run timestamp", Format(Now, "yyyy-mm-dd hh:nn:ss")
    AddStatRow wsStats, r, "Mailbox", IIf(Len(SHARED_MAILBOX) = 0, "<your own>", SHARED_MAILBOX)
    AddStatRow wsStats, r, "Folder(s)", folderDesc
    AddStatRow wsStats, r, "Scan window", windowDesc
    AddStatRow wsStats, r, "Mode", "READ-ONLY (no mailbox changes)"
    AddStatRow wsStats, r, "Rule", "in-hours = Paris team OR India team on the clock"
    AddStatRow wsStats, r, "India shift (IST)", _
        Format(INDIA_SHIFT_START_IST_MIN \ 60, "00") & ":" & Format(INDIA_SHIFT_START_IST_MIN Mod 60, "00") & _
        " - " & Format(INDIA_SHIFT_END_IST_MIN \ 60, "00") & ":" & Format(INDIA_SHIFT_END_IST_MIN Mod 60, "00")
    AddStatRow wsStats, r, "Output file", outPath
    r = r + 1

    SectionHeader wsStats, r, "Totals"
    AddStatRow wsStats, r, "Total mails scanned", total
    AddStatRow wsStats, r, "In-hours (IsOffHours=FALSE)", total - offHoursCount
    AddStatRow wsStats, r, "Off-hours (IsOffHours=TRUE)", offHoursCount
    Dim pctRow As Long: pctRow = r
    AddStatRow wsStats, r, "Off-hours share", IIf(total > 0, offHoursCount / total, 0)
    wsStats.Cells(pctRow, 2).NumberFormat = "0.0%"
    r = r + 1

    SectionHeader wsStats, r, "In-hours coverage (categories overlap)"
    AddStatRow wsStats, r, "Covered by Paris team", countParisHours
    AddStatRow wsStats, r, "Covered by India team", countIndiaHours
    AddStatRow wsStats, r, "India-only (Paris off, India on)", countIndiaRescued
    r = r + 1

    SectionHeader wsStats, r, "Holiday flags seen (info only, overlap)"
    AddStatRow wsStats, r, "France holiday", countHolidayFR
    AddStatRow wsStats, r, "Poland holiday", countHolidayPL
    AddStatRow wsStats, r, "India (Mumbai) holiday", countHolidayIN

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
           "Folder:  " & folderDesc & vbCrLf & _
           "Window:  " & windowDesc & vbCrLf & _
           "Mode:    READ-ONLY (no mailbox changes)" & vbCrLf & _
           "Rule:    in-hours = Paris team OR India team on the clock" & vbCrLf & vbCrLf & _
           "Scanned " & total & " mails - all written to the Emails sheet." & vbCrLf & _
           "Off-hours (IsOffHours=TRUE): " & offHoursCount & vbCrLf & vbCrLf & _
           "In-hours coverage breakdown (categories overlap):" & vbCrLf & _
           "  Covered by Paris team:   " & countParisHours & vbCrLf & _
           "  Covered by India team:   " & countIndiaHours & vbCrLf & _
           "  India-only (Paris off):  " & countIndiaRescued & "  <- counted in-hours here," & vbCrLf & _
           "                              would be off-hours in the Paris-only script" & vbCrLf & vbCrLf & _
           "Holiday flags seen (info):" & vbCrLf & _
           "  France holiday:          " & countHolidayFR & vbCrLf & _
           "  Poland holiday:          " & countHolidayPL & vbCrLf & _
           "  India (Mumbai) holiday:  " & countHolidayIN & vbCrLf & vbCrLf & _
           "Excel is now open at:" & vbCrLf & outPath & vbCrLf & vbCrLf & _
           "Paste your app users into the 'Users' tab (column A). The IsAppUser " & _
           "column on the Emails tab will then show TRUE/FALSE per row.", _
           vbInformation, "Off-Hours Email Finder (India in-hours)"
End Sub


' Diagnostic helper - print top-level mailbox display names to the Immediate
' window (Ctrl+G). Read-only. Use when you have multiple Inboxes and need the
' exact name for SHARED_MAILBOX.
Public Sub ListMailboxes()
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    Debug.Print "=== Top-level mailboxes in your Outlook profile ==="
    Debug.Print "Copy a name BETWEEN the quotes EXACTLY into SHARED_MAILBOX."
    Debug.Print ""

    Dim i As Long
    For i = 1 To ns.folders.Count
        Dim root As Outlook.folder
        Set root = ns.folders(i)
        Dim inboxCount As String
        inboxCount = "?"
        On Error Resume Next
        inboxCount = CStr(root.folders("Inbox").items.Count)
        On Error GoTo 0
        Debug.Print "  " & i & ".  """ & root.Name & """    (Inbox items: " & inboxCount & ")"
    Next i

    Debug.Print ""
    Debug.Print "=== End ==="
    MsgBox "Mailbox list printed to the Immediate window (Ctrl+G in the VBA editor).", _
        vbInformation, "ListMailboxes"
End Sub


' Write one "label | value" row to the Stats sheet and advance the row counter.
Private Sub AddStatRow(ByVal ws As Object, ByRef r As Long, ByVal label As String, ByVal value As Variant)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Value = value
    r = r + 1
End Sub

' Write a bold section header to the Stats sheet and advance the row counter.
Private Sub SectionHeader(ByVal ws As Object, ByRef r As Long, ByVal title As String)
    ws.Cells(r, 1).Value = title
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
End Sub


' Generic membership check against a YYYY-MM-DD CSV list.
Private Function IsInDateList(ByVal t As Date, ByVal listCsv As String) As Boolean
    If Len(listCsv) = 0 Then Exit Function
    Dim key As String
    key = Format(t, "yyyy-mm-dd")
    Dim parts() As String
    parts = Split(listCsv, ",")
    Dim i As Long
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


' True if the date falls within EU summer time (CEST). EU rule: last Sunday of
' March -> last Sunday of October. Date-granularity is enough here.
Private Function IsSummerTime(ByVal t As Date) As Boolean
    Dim y As Integer
    y = Year(t)
    Dim startDST As Date, endDST As Date
    startDST = LastSundayOfMonth(y, 3)
    endDST = LastSundayOfMonth(y, 10)
    IsSummerTime = (DateValue(t) >= startDST And DateValue(t) < endDST)
End Function


' Last Sunday of the given month/year.
Private Function LastSundayOfMonth(ByVal y As Integer, ByVal m As Integer) As Date
    Dim lastDay As Date
    lastDay = DateSerial(y, m + 1, 0)
    Dim dow As Integer
    dow = Weekday(lastDay, vbMonday)
    LastSundayOfMonth = lastDay - (dow Mod 7)
End Function


' Walk the folder tree, populating parallel folder/path Collections.
' Read-only: only enumerates folder names.
Private Sub CollectFolders(ByVal root As Outlook.folder, _
                           ByVal pathSoFar As String, _
                           ByRef folderColl As Collection, _
                           ByRef pathColl As Collection, _
                           ByVal includeSubfolders As Boolean)
    Dim rootPath As String
    If Len(pathSoFar) = 0 Then
        rootPath = root.Name
    Else
        rootPath = pathSoFar & "/" & root.Name
    End If
    folderColl.Add root
    pathColl.Add rootPath

    If Not includeSubfolders Then Exit Sub

    Dim child As Outlook.folder
    For Each child In root.folders
        CollectFolders child, rootPath, folderColl, pathColl, True
    Next child
End Sub


' Resolve the target folder (own mailbox / shared by SMTP / shared by name).
Private Function ResolveFolder(ByVal ns As Outlook.NameSpace) As Outlook.folder
    On Error Resume Next

    If Len(SHARED_MAILBOX) = 0 Then
        If StrComp(FOLDER_NAME, "Sent Items", vbTextCompare) = 0 Then
            Set ResolveFolder = ns.GetDefaultFolder(olFolderSentMail)
        ElseIf StrComp(FOLDER_NAME, "Inbox", vbTextCompare) = 0 Then
            Set ResolveFolder = ns.GetDefaultFolder(olFolderInbox)
        Else
            Set ResolveFolder = ns.GetDefaultFolder(olFolderInbox).folders(FOLDER_NAME)
        End If
        Exit Function
    End If

    If InStr(SHARED_MAILBOX, "@") > 0 Then
        Dim recip As Outlook.Recipient
        Set recip = ns.CreateRecipient(SHARED_MAILBOX)
        recip.Resolve
        If Not recip.Resolved Then Exit Function

        Dim sharedRoot As Outlook.folder
        If StrComp(FOLDER_NAME, "Sent Items", vbTextCompare) = 0 Then
            Set ResolveFolder = ns.GetSharedDefaultFolder(recip, olFolderSentMail)
        Else
            Set sharedRoot = ns.GetSharedDefaultFolder(recip, olFolderInbox)
            If StrComp(FOLDER_NAME, "Inbox", vbTextCompare) = 0 Then
                Set ResolveFolder = sharedRoot
            Else
                Set ResolveFolder = sharedRoot.folders(FOLDER_NAME)
            End If
        End If
        Exit Function
    End If

    Dim store As Outlook.folder
    Set store = ns.folders(SHARED_MAILBOX)
    If store Is Nothing Then Exit Function
    Set ResolveFolder = store.folders(FOLDER_NAME)
End Function


' SMTP address of the sender, falling back to display name.
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
    SafeSenderAddress = addr
End Function


' Build the full XLSX path. Auto-creates the output directory if missing.
Private Function ResolveOutputPath() As String
    Dim outDir As String
    If Len(OUTPUT_DIR) > 0 Then
        outDir = OUTPUT_DIR
    Else
        outDir = GetDocumentsPath() & "\OffHoursEmails"
    End If
    If Right(outDir, 1) <> "\" Then outDir = outDir & "\"

    On Error Resume Next
    If Len(Dir(outDir, vbDirectory)) = 0 Then MkDir outDir
    On Error GoTo 0

    If Len(Dir(outDir, vbDirectory)) = 0 Then
        MsgBox "Cannot create output directory:" & vbCrLf & outDir, vbCritical, "Off-Hours Email Finder (India in-hours)"
        Exit Function
    End If

    If ARCHIVE_FILES Then
        Dim stamp As String
        stamp = Format(Now, "yyyy-mm-dd") & "_" & Format(Now, "hhnn")
        If Len(SHARED_MAILBOX) > 0 Then
            ResolveOutputPath = outDir & FILENAME_PREFIX & "-" & _
                SanitizeForFilename(SHARED_MAILBOX) & "-" & stamp & ".xlsx"
        Else
            ResolveOutputPath = outDir & FILENAME_PREFIX & "-" & stamp & ".xlsx"
        End If
    Else
        ResolveOutputPath = outDir & FILENAME_PREFIX & ".xlsx"
    End If
End Function


' Make a string safe to use in a filename.
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


' Resolve the user's real Documents folder (handles OneDrive redirection).
Private Function GetDocumentsPath() As String
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    If Not wsh Is Nothing Then
        GetDocumentsPath = wsh.SpecialFolders("MyDocuments")
    End If
    On Error GoTo 0
    If Len(GetDocumentsPath) = 0 Then
        GetDocumentsPath = Environ("USERPROFILE") & "\Documents"
    End If
End Function
