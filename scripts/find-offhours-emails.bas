Attribute VB_Name = "OffHoursEmails"
' ============================================================================
'  Off-Hours Email Finder for Outlook
'  Project: p5-outlook-offhours-emails (APS-llm-tools)
'
'  Iterates a chosen Outlook folder (and optionally its subfolders) and
'  writes an XLSX with two sheets:
'
'  Sheet "Emails" - one row per MailItem in the scan window with timing
'      classification (receive time, hour-of-day, weekday, weekend flag,
'      per-calendar holiday flags FR / PL / IN, IsOffHours union flag),
'      sender ("From"), Subject, EntryID, the source FolderPath, and a
'      built-in IsAppUser formula joining against the Users tab.
'
'  Sheet "Users" - empty tab for the analyst to paste the application's
'      user-list emails (one per row in column A). The IsAppUser formula
'      on the Emails sheet recalculates automatically against this list.
'
'  Subfolder traversal is OFF by default - the macro scans only FOLDER_NAME
'  itself. Set INCLUDE_SUBFOLDERS = True to recurse into every descendant
'  folder (e.g. Archive, Project subfolders). Each folder's items are still
'  filtered server-side via Items.Restrict, so per-folder cost stays low.
'
'  Sender is captured as a column but NOT pre-classified by the macro.
'  The IsAppUser formula does the join against your user database list.
'
'  READ-ONLY. NO WRITE CAPABILITY.
'  -----------------------------------------------------------------
'  This macro contains NO code that writes to any MailItem or any
'  Outlook / Exchange object. There is no .Save, no Categories
'  assignment, no UnRead change, no Move / Delete / Reply / Forward /
'  Send anywhere in this module. It only reads:
'      ReceivedTime, Subject, Sender / SenderEmailAddress / SenderName,
'      EntryID.
'  This is a structural property of the source - not a runtime flag.
'  Safe to run against shared / project / team mailboxes; no other
'  user with access sees any change.
'
'  The ONLY thing this macro writes is the local XLSX under OUTPUT_DIR
'  (default: <Documents>\OffHoursEmails\).
'  -----------------------------------------------------------------
'
'  How to use:
'    1. Outlook -> Alt+F11 -> File -> Import File... -> pick this .bas
'    2. Adjust the CONFIG constants below (especially OUTPUT_DIR,
'       SHARED_MAILBOX if scanning a project inbox, and the HOLIDAYS_*_CSV lists)
'    3. Run macro: FindOffHoursEmails (F5)
' ============================================================================

Option Explicit

' -------------------- CONFIG (edit these) --------------------
' Working hours are interpreted in your local time zone. Outlook stores
' MailItem.ReceivedTime in the local time of the machine running Outlook -
' for a Paris user that is CET/CEST automatically; no conversion is done
' or needed. An email is in-hours when WORK_START_HOUR <= Hour(ReceivedTime)
' < WORK_END_HOUR.
Private Const WORK_START_HOUR As Integer = 9           ' inclusive, local time (09:00)
Private Const WORK_END_HOUR As Integer = 18            ' exclusive, local time (18:00 = 6 PM)
Private Const TREAT_SATURDAY_AS_OFFHOURS As Boolean = True
Private Const TREAT_SUNDAY_AS_OFFHOURS As Boolean = True

' India team support coverage end time on a Polish holiday (Paris local time).
' Mumbai works 18:00 IST. IST is UTC+5:30 fixed (no DST). Paris is UTC+1
' (CET, winter) or UTC+2 (CEST, summer). So 18:00 IST = 13:30 Paris winter,
' 14:30 Paris summer. The macro auto-detects DST via IsSummerTime().
' Stored as minutes-since-midnight so 13:30 / 14:30 are exact.
Private Const INDIA_END_SUMMER_MIN As Long = 14 * 60 + 30   ' 14:30 = 870
Private Const INDIA_END_WINTER_MIN As Long = 13 * 60 + 30   ' 13:30 = 810

' Public holidays (off-hours regardless of weekday). Comma-separated YYYY-MM-DD.
' Three independent calendars - the macro flags an email's date with each
' calendar separately (IsHolidayFR / IsHolidayPL / IsHolidayIN columns) and
' the summary breaks counts down per calendar. The off-hours decision is
' the UNION: an email is off-hours if its date is a holiday in ANY of the
' three lists. Set any list to "" to disable that calendar entirely.

' France 2025 + 2026 - national work-free days. 2025 entries kept for the
' SCAN_LAST_N_DAYS = 365 lookback into the previous year.
'   2025
'   01 Jan  - New Year's Day
'   21 Apr  - Easter Monday
'   01 May  - Labour Day
'   08 May  - Victory in Europe Day
'   29 May  - Ascension Day
'   09 Jun  - Whit Monday
'   14 Jul  - Bastille Day
'   15 Aug  - Assumption
'   01 Nov  - All Saints' Day
'   11 Nov  - Armistice Day
'   25 Dec  - Christmas Day
'   2026
'   01 Jan  - New Year's Day
'   06 Apr  - Easter Monday
'   01 May  - Labour Day
'   08 May  - Victory in Europe Day
'   14 May  - Ascension Day
'   25 May  - Whit Monday
'   14 Jul  - Bastille Day
'   15 Aug  - Assumption
'   01 Nov  - All Saints' Day
'   11 Nov  - Armistice Day
'   25 Dec  - Christmas Day
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

' Poland 2025 + 2026 - work-free days. Includes Sundays formally listed as
' public holidays (Easter Sunday, Pentecost Sunday). Christmas Eve became a
' public holiday in 2025. 2025 entries kept for the 365-day lookback.
'   2025
'   01 Jan  - New Year's Day
'   06 Jan  - Epiphany
'   20 Apr  - Easter Sunday
'   21 Apr  - Easter Monday
'   01 May  - Labour Day
'   03 May  - Constitution Day
'   08 Jun  - Pentecost Sunday
'   19 Jun  - Corpus Christi
'   15 Aug  - Assumption / Polish Armed Forces Day
'   01 Nov  - All Saints' Day
'   11 Nov  - Independence Day
'   24 Dec  - Christmas Eve (public holiday since 2025)
'   25 Dec  - Christmas Day
'   26 Dec  - Second Day of Christmas
'   2026
'   01 Jan  - New Year's Day
'   06 Jan  - Epiphany
'   05 Apr  - Easter Sunday
'   06 Apr  - Easter Monday
'   01 May  - Labour Day
'   03 May  - Constitution Day
'   24 May  - Pentecost Sunday
'   04 Jun  - Corpus Christi
'   15 Aug  - Assumption / Polish Armed Forces Day
'   01 Nov  - All Saints' Day
'   11 Nov  - Independence Day
'   24 Dec  - Christmas Eve
'   25 Dec  - Christmas Day
'   26 Dec  - Second Day of Christmas
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

' India - Mumbai (Maharashtra) 2025 + 2026 bank holidays
' VERIFY before relying on stats: festival dates depend on lunar / regional
' calendars and shift each year. The authoritative list is published yearly
' by the RBI for Maharashtra. Replace this CSV with your HR's official list.
' 2025 entries are kept here because SCAN_LAST_N_DAYS = 365 looks back into
' the previous calendar year.
'   2025
'   26 Jan  - Republic Day
'   19 Feb  - Chhatrapati Shivaji Maharaj Jayanti
'   14 Mar  - Holi (Dhulivandan)
'   31 Mar  - Eid-ul-Fitr (Ramzan Id)
'   10 Apr  - Mahavir Jayanti
'   14 Apr  - Dr. Ambedkar Jayanti
'   18 Apr  - Good Friday
'   01 May  - Maharashtra Day
'   12 May  - Buddha Purnima
'   07 Jun  - Bakri Id (Eid-ul-Adha)
'   06 Jul  - Muharram
'   15 Aug  - Independence Day / Parsi New Year
'   27 Aug  - Ganesh Chaturthi
'   02 Oct  - Gandhi Jayanti / Dussehra
'   21 Oct  - Diwali (Lakshmi Pujan)
'   22 Oct  - Diwali Padwa / Balipratipada
'   05 Nov  - Guru Nanak Jayanti
'   25 Dec  - Christmas Day
'   2026
'   26 Jan  - Republic Day
'   04 Mar  - Holi
'   19 Mar  - Gudi Padwa (Maharashtra New Year)
'   03 Apr  - Good Friday
'   14 Apr  - Dr. Ambedkar Jayanti
'   01 May  - Maharashtra Day (overlaps with French Labour Day)
'   15 Aug  - Independence Day
'   14 Sep  - Ganesh Chaturthi
'   02 Oct  - Gandhi Jayanti
'   20 Oct  - Dussehra
'   08 Nov  - Diwali (Lakshmi Pujan)
'   09 Nov  - Diwali Padwa / Bhai Dooj
'   25 Dec  - Christmas Day
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
'   "Project Inbox Name"     = a shared mailbox by display name (as it appears
'                              in the Outlook folder pane root)
'   "project@company.com"    = a shared mailbox by SMTP address
'                              (resolved via NameSpace.CreateRecipient)
Private Const SHARED_MAILBOX As String = ""

' Folder name within the resolved mailbox: "Inbox", "Sent Items", or a direct
' subfolder name under Inbox.
Private Const FOLDER_NAME As String = "Inbox"

' Recursive subfolder traversal.
'   False (default) - scan only FOLDER_NAME itself; mails moved out of Inbox
'                     into subfolders (Archive, Project A, etc.) are NOT seen.
'   True            - scan FOLDER_NAME plus every descendant folder under it.
'                     Items.Restrict is applied per folder, so each folder's
'                     scan stays index-side and fast.
' When True, the Emails sheet's "FolderPath" column shows the relative folder
' path of each row (e.g. "Inbox", "Inbox/Archive", "Inbox/Project A/2025"),
' so you can group / pivot by folder. When False, every row's FolderPath
' equals the root folder name.
Private Const INCLUDE_SUBFOLDERS As Boolean = False

' --- Scan window ---
' Two ways to set the window. Pick ONE pattern.
'
' Pattern A - explicit date range (recommended for monthly extraction):
'     SCAN_FROM_DATE = "2025-04-01"  -> first day, 00:00
'     SCAN_TO_DATE   = "2025-04-30"  -> last day, 23:59:59
'   For other months / quarters / years, just change the two dates.
'
' Pattern B - last N days from today (fallback when SCAN_FROM_DATE is empty):
'     SCAN_FROM_DATE = ""
'     SCAN_TO_DATE   = ""
'     SCAN_LAST_N_DAYS = 30
'
' The macro uses Outlook's Items.Restrict() to apply the filter server/index
' side, so a small window (e.g. one month) is fast even if it sits years
' back in a large mailbox - only the matching items are touched.
'
' Date format: "YYYY-MM-DD". Leave a value empty ("") to disable that bound.
Private Const SCAN_FROM_DATE As String = ""
Private Const SCAN_TO_DATE As String = ""
Private Const SCAN_LAST_N_DAYS As Long = 30

' Output location. Each run writes a new timestamped XLSX so the history of
' every scan is preserved (set ARCHIVE_FILES = False to overwrite a single
' file instead). Default directory:
'   <Documents>\OffHoursEmails\
' where <Documents> resolves to your real Documents folder via
' WScript.Shell.SpecialFolders("MyDocuments") - this handles OneDrive /
' corporate folder-redirection automatically and falls back to
' %USERPROFILE%\Documents if WScript.Shell is unavailable.
' The directory is auto-created on first run; no placeholder needed.
'
' Override OUTPUT_DIR with an absolute path (e.g. "C:\Reports\OffHours\")
' if you want the file somewhere else.
Private Const OUTPUT_DIR As String = ""           ' "" = default to <Documents>\OffHoursEmails\
Private Const ARCHIVE_FILES As Boolean = True      ' True = timestamped filenames; False = single overwriting file
Private Const FILENAME_PREFIX As String = "offhours-emails"
' -------------------------------------------------------------


Public Sub FindOffHoursEmails()
    Dim ns As Outlook.NameSpace
    Dim folder As Outlook.folder
    Dim itm As Object
    Dim total As Long, offHoursCount As Long
    Dim countHolidayFR As Long, countHolidayPL As Long, countHolidayIN As Long
    Dim countWeekend As Long, countOutsideHours As Long

    Set ns = Application.GetNamespace("MAPI")
    Set folder = ResolveFolder(ns)
    If folder Is Nothing Then
        MsgBox "Folder not found. Mailbox=""" & SHARED_MAILBOX & """, Folder=""" & FOLDER_NAME & """", vbExclamation
        Exit Sub
    End If

    Dim outPath As String
    outPath = ResolveOutputPath()
    If Len(outPath) = 0 Then Exit Sub   ' ResolveOutputPath has already shown an error

    ' Resolve the scan window: explicit date range first, fallback to N days back.
    Dim fromDt As Date, toDt As Date
    Dim windowDesc As String

    If Len(SCAN_FROM_DATE) > 0 Then
        On Error Resume Next
        fromDt = CDate(SCAN_FROM_DATE)
        If Err.Number <> 0 Then
            MsgBox "SCAN_FROM_DATE could not be parsed: """ & SCAN_FROM_DATE & """" & vbCrLf & _
                   "Use YYYY-MM-DD (e.g. 2025-04-01).", vbExclamation, "Off-Hours Email Finder"
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
                   "Use YYYY-MM-DD (e.g. 2025-04-30).", vbExclamation, "Off-Hours Email Finder"
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    End If

    ' Build Outlook restriction filter. Server/index-side filtering is far
    ' faster than iterate-and-skip - especially for narrow windows that sit
    ' months or years back in a large mailbox.
    Dim restriction As String
    If fromDt > 0 Then
        restriction = "[ReceivedTime] >= '" & Format(fromDt, "ddddd h:nn AMPM") & "'"
    End If
    If toDt > 0 Then
        If Len(restriction) > 0 Then restriction = restriction & " AND "
        restriction = restriction & "[ReceivedTime] <= '" & Format(toDt, "ddddd h:nn AMPM") & "'"
    End If

    ' Human-readable description of the resolved window for the summary dialog
    If fromDt > 0 And toDt > 0 Then
        windowDesc = Format(fromDt, "yyyy-mm-dd") & " to " & Format(toDt, "yyyy-mm-dd")
    ElseIf fromDt > 0 Then
        windowDesc = "from " & Format(fromDt, "yyyy-mm-dd") & " to now"
    ElseIf toDt > 0 Then
        windowDesc = "everything up to " & Format(toDt, "yyyy-mm-dd")
    Else
        windowDesc = "whole folder (no date filter)"
    End If

    ' Build the list of folders to scan: just the resolved root, or root +
    ' every descendant if INCLUDE_SUBFOLDERS = True. Items.Restrict is applied
    ' per folder, so each folder's scan stays index-side and fast.
    Dim folderList As Collection, pathList As Collection
    Set folderList = New Collection
    Set pathList = New Collection
    CollectFolders folder, "", folderList, pathList, INCLUDE_SUBFOLDERS

    ' Sum items.Count across all folders to size the buffer (upper bound -
    ' folder.items.Count is fast metadata, unlike Restrict.Count which forces
    ' filter evaluation).
    Dim maxRows As Long
    Dim fIdx As Long
    maxRows = 0
    For fIdx = 1 To folderList.Count
        Dim fTmp As Outlook.folder
        Set fTmp = folderList(fIdx)
        maxRows = maxRows + fTmp.items.Count
    Next fIdx
    If maxRows < 1 Then maxRows = 1

    ' Pre-allocate two buffers (split by column-I formula). Final sheet layout:
    '   A:H  = leftData  (timing block: ReceivedTime ... IsOffHours)
    '   I    = IsAppUser formula
    '   J:M  = rightData (From, Subject, EntryID, FolderPath)
    Dim leftData() As Variant   ' 8 cols
    Dim rightData() As Variant  ' 4 cols (FolderPath added at index 4)
    ReDim leftData(1 To maxRows, 1 To 8)
    ReDim rightData(1 To maxRows, 1 To 4)
    Dim row As Long
    row = 0

    ' Iterate every folder in the list. Items.Restrict is applied per folder.
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
        folderScan.Sort "[ReceivedTime]", True   ' newest first within this folder

        For Each itm In folderScan
            If TypeOf itm Is Outlook.MailItem Then
                Dim m As Outlook.MailItem
                Set m = itm
                total = total + 1

                ' Compute timing flags for this email
                Dim hFR As Boolean, hPL As Boolean, hIN As Boolean
                Dim isWknd As Boolean, hourLocal As Integer
                Dim minuteOfDay As Long, outsideOffice As Boolean, isOff As Boolean
                hFR = IsHolidayFR(m.ReceivedTime)
                hPL = IsHolidayPL(m.ReceivedTime)
                hIN = IsHolidayIN(m.ReceivedTime)
                isWknd = IsWeekendDay(m.ReceivedTime)
                hourLocal = Hour(m.ReceivedTime)
                minuteOfDay = hourLocal * 60 + Minute(m.ReceivedTime)
                outsideOffice = (minuteOfDay < WORK_START_HOUR * 60 Or _
                                 minuteOfDay >= WORK_END_HOUR * 60)

                ' IsOffHours rule (see PLAN.md §1):
                '   - Saturday/Sunday => off-hours all day
                '   - Outside 09:00-18:00 Paris local => off-hours
                '   - PL holiday during office hours, AND (IN holiday OR past
                '     India coverage end - 14:30 summer / 13:30 winter)
                ' France and Mumbai holiday flags are tracked but do NOT drive
                ' isOff; Mumbai only enters as the India-team-also-off check.
                Dim indiaEndMin As Long
                If IsSummerTime(m.ReceivedTime) Then
                    indiaEndMin = INDIA_END_SUMMER_MIN
                Else
                    indiaEndMin = INDIA_END_WINTER_MIN
                End If
                isOff = (isWknd And TREAT_SATURDAY_AS_OFFHOURS And Weekday(m.ReceivedTime, vbMonday) = 6) Or _
                        (isWknd And TREAT_SUNDAY_AS_OFFHOURS And Weekday(m.ReceivedTime, vbMonday) = 7) Or _
                        outsideOffice Or _
                        (hPL And (hIN Or minuteOfDay >= indiaEndMin))

                ' Tallies (overlap intentional - one mail can be e.g. weekend AND holiday)
                If isOff Then offHoursCount = offHoursCount + 1
                If hFR Then countHolidayFR = countHolidayFR + 1
                If hPL Then countHolidayPL = countHolidayPL + 1
                If hIN Then countHolidayIN = countHolidayIN + 1
                If isWknd Then countWeekend = countWeekend + 1
                If outsideOffice Then countOutsideHours = countOutsideHours + 1

                row = row + 1
                leftData(row, 1) = Format(m.ReceivedTime, "yyyy-mm-dd hh:nn:ss")
                leftData(row, 2) = hourLocal
                leftData(row, 3) = WeekdayName(Weekday(m.ReceivedTime, vbMonday), True)
                leftData(row, 4) = isWknd
                leftData(row, 5) = hFR
                leftData(row, 6) = hPL
                leftData(row, 7) = hIN
                leftData(row, 8) = isOff
                rightData(row, 1) = SafeSenderAddress(m)
                rightData(row, 2) = m.Subject
                rightData(row, 3) = m.EntryID
                rightData(row, 4) = curPath
            End If
        Next itm
    Next fIdx

    If row = 0 Then
        MsgBox "No mails in the scan window. No file written.", vbInformation, "Off-Hours Email Finder"
        Exit Sub
    End If

    ' --- Write XLSX via Excel automation (no add-in dependency) ---
    Dim xl As Object
    On Error Resume Next
    Set xl = CreateObject("Excel.Application")
    On Error GoTo 0
    If xl Is Nothing Then
        MsgBox "Could not start Excel. The macro requires Excel to write the XLSX. " & _
               "Confirm Excel is installed and try again.", vbCritical, "Off-Hours Email Finder"
        Exit Sub
    End If

    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False

    Dim wb As Object
    Set wb = xl.Workbooks.Add

    ' Reduce default sheet count to 1
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop

    ' --- Sheet 1: "Emails" ---
    Dim ws As Object
    Set ws = wb.Worksheets(1)
    ws.Name = "Emails"

    Dim hdrs As Variant
    hdrs = Array("ReceivedTime", "HourLocal", "Weekday", "IsWeekend", _
                 "IsHolidayFR", "IsHolidayPL", "IsHolidayIN", "IsOffHours", _
                 "IsAppUser", "From", "Subject", "EntryID", "FolderPath")
    Dim hi As Long
    For hi = 0 To UBound(hdrs)
        ws.Cells(1, hi + 1).Value = hdrs(hi)
    Next hi
    ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(hdrs) + 1)).Font.Bold = True

    ' Bulk-write the data blocks. Sheet layout:
    '   A:H = leftData (timing flags)
    '   I   = IsAppUser formula (filled below)
    '   J:M = rightData (From, Subject, EntryID, FolderPath)
    ws.Range(ws.Cells(2, 1), ws.Cells(maxRows + 1, 8)).Value = leftData
    ws.Range(ws.Cells(2, 10), ws.Cells(maxRows + 1, 13)).Value = rightData
    If row < maxRows Then
        ws.Range(ws.Cells(row + 2, 1), ws.Cells(maxRows + 1, 8)).ClearContents
        ws.Range(ws.Cells(row + 2, 10), ws.Cells(maxRows + 1, 13)).ClearContents
    End If

    ' IsAppUser formula in col I for populated rows only.
    ' Returns TRUE/FALSE based on whether the From value (col J) appears in
    ' Users!A:A. Empty Users tab = FALSE for everything (correct default).
    ws.Range(ws.Cells(2, 9), ws.Cells(row + 1, 9)).Formula = _
        "=IF(J2="""",FALSE,COUNTIF(Users!A:A,J2)>0)"

    ' AutoFilter on the full data range
    ws.Range(ws.Cells(1, 1), ws.Cells(row + 1, 13)).AutoFilter

    ws.Columns("A:M").AutoFit

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
    wsStats.Cells(1, 1).Value = "Off-Hours Email Finder - Run Statistics"
    wsStats.Cells(1, 1).Font.Bold = True
    wsStats.Cells(1, 1).Font.Size = 12

    Dim r As Long
    r = 3
    AddStatRow wsStats, r, "Run timestamp", Format(Now, "yyyy-mm-dd hh:nn:ss")
    AddStatRow wsStats, r, "Mailbox", IIf(Len(SHARED_MAILBOX) = 0, "<your own>", SHARED_MAILBOX)
    AddStatRow wsStats, r, "Folder(s)", folderDesc
    AddStatRow wsStats, r, "Scan window", windowDesc
    AddStatRow wsStats, r, "Mode", "READ-ONLY (no mailbox changes)"
    AddStatRow wsStats, r, "Rule", "off-hours = weekend OR outside " & WORK_START_HOUR & _
        "-" & WORK_END_HOUR & " Paris OR (Polish holiday AND India also off/past coverage end)"
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

    SectionHeader wsStats, r, "Off-hours breakdown (categories overlap)"
    AddStatRow wsStats, r, "Weekend", countWeekend
    AddStatRow wsStats, r, "Outside " & WORK_START_HOUR & "-" & WORK_END_HOUR, countOutsideHours
    r = r + 1

    SectionHeader wsStats, r, "Holiday flags seen (info; only PL drives off-hours)"
    AddStatRow wsStats, r, "France holiday", countHolidayFR
    AddStatRow wsStats, r, "Poland holiday", countHolidayPL
    AddStatRow wsStats, r, "India (Mumbai) holiday", countHolidayIN

    wsStats.Columns("A:B").AutoFit

    ' Re-activate Emails sheet, freeze top row, position cursor at A2
    ws.Activate
    xl.ActiveWindow.SplitRow = 1
    xl.ActiveWindow.FreezePanes = True
    ws.Range("A2").Select

    ' Save as XLSX (51 = xlOpenXMLWorkbook)
    wb.SaveAs outPath, 51

    xl.ScreenUpdating = True
    xl.DisplayAlerts = True
    xl.Visible = True   ' show the workbook to the user; do NOT quit Excel

    MsgBox "Mailbox: " & IIf(Len(SHARED_MAILBOX) = 0, "<your own>", SHARED_MAILBOX) & vbCrLf & _
           "Folder:  " & folderDesc & vbCrLf & _
           "Window:  " & windowDesc & vbCrLf & _
           "Mode:    READ-ONLY (no mailbox changes)" & vbCrLf & vbCrLf & _
           "Scanned " & total & " mails - all written to the Emails sheet." & vbCrLf & _
           "Off-hours (IsOffHours=TRUE): " & offHoursCount & vbCrLf & vbCrLf & _
           "Of the off-hours mails (categories overlap):" & vbCrLf & _
           "  Weekend:                 " & countWeekend & vbCrLf & _
           "  Outside " & WORK_START_HOUR & "-" & WORK_END_HOUR & ":              " & countOutsideHours & vbCrLf & _
           "  France holiday:          " & countHolidayFR & vbCrLf & _
           "  Poland holiday:          " & countHolidayPL & vbCrLf & _
           "  India (Mumbai) holiday:  " & countHolidayIN & vbCrLf & vbCrLf & _
           "Excel is now open at:" & vbCrLf & outPath & vbCrLf & vbCrLf & _
           "Paste your app users into the 'Users' tab (column A). The IsAppUser " & _
           "column on the Emails tab will then show TRUE/FALSE per row.", _
           vbInformation, "Off-Hours Email Finder"
End Sub


' Diagnostic helper. Run this ONCE (Alt+F11 -> click in this sub -> F5) when
' you have multiple Inboxes (your own + shared/team mailboxes) and you need
' to know the exact display name to put in SHARED_MAILBOX. Output goes to
' the Immediate window (Ctrl+G in the VBA editor).
'
' Read-only: only enumerates root folder display names and Inbox item counts.
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


Private Function IsWeekendDay(ByVal t As Date) As Boolean
    Dim dow As Integer
    dow = Weekday(t, vbMonday)
    IsWeekendDay = (dow = 6 Or dow = 7)
End Function


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


' True if the date falls within EU summer time (CEST). EU rule: summer time
' runs from the last Sunday of March (clocks forward 02:00 -> 03:00) to the
' last Sunday of October (clocks back 03:00 -> 02:00). Date-granularity is
' enough for our purposes - the 1-hour window at the switchover Sundays is
' not meaningful for off-hours classification.
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
    lastDay = DateSerial(y, m + 1, 0)   ' day 0 of month m+1 = last day of m
    Dim dow As Integer
    dow = Weekday(lastDay, vbMonday)    ' 1=Mon..7=Sun
    LastSundayOfMonth = lastDay - (dow Mod 7)
End Function


' Walk the folder tree starting from a root, populating two parallel
' Collections:
'   folderColl - the Outlook.folder objects, in depth-first order
'   pathColl   - matching relative path strings (e.g. "Inbox", "Inbox/Archive")
' If includeSubfolders is False, only the root is added.
'
' The path of the root is just its own name; children are "<rootName>/<child>",
' grandchildren are "<rootName>/<child>/<grandchild>", and so on. The path is
' the FolderPath column value written for every mail under that folder.
'
' Read-only: only enumerates folder names; no items are touched here.
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


' Resolve the target folder.
'   - SHARED_MAILBOX = ""             -> user's own default mailbox
'   - SHARED_MAILBOX contains "@"     -> shared mailbox resolved by SMTP via Recipient
'   - else                            -> shared mailbox already mounted, resolved
'                                        by display name in the Outlook root folder list
Private Function ResolveFolder(ByVal ns As Outlook.NameSpace) As Outlook.folder
    On Error Resume Next

    If Len(SHARED_MAILBOX) = 0 Then
        ' User's own mailbox
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
        ' Shared mailbox by SMTP - resolve via Recipient
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

    ' Shared mailbox already mounted in the folder pane - resolve by display name
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
' Returns "" on failure (after showing an error to the user).
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
        MsgBox "Cannot create output directory:" & vbCrLf & outDir, vbCritical, "Off-Hours Email Finder"
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


' Make a string safe to use in a filename. Keeps letters, digits, dash, and
' underscore; spaces become underscores; everything else is dropped.
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


' Resolve the user's real Documents folder. Tries WScript.Shell first (handles
' OneDrive / corporate folder-redirection); falls back to %USERPROFILE%\Documents.
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
