Attribute VB_Name = "TimingColumns"
' ============================================================================
'  Excel Timing Classifier
'  Project: p5-outlook-offhours-emails (APS-llm-tools)
'
'  Generic helper that adds timing-classification columns to ANY Excel sheet
'  next to a date column you specify. Same working-hours / holiday rules
'  as the Outlook macro (find-offhours-emails.bas). Use it for ServiceNow
'  ticket exports (OpeningDate), Jira CSV-to-XLSX (Created), HR exports of
'  login times - any sheet where one column carries a timestamp.
'
'  How it works:
'    1. You open the .xlsx file in Excel.
'    2. Alt+F11 -> File -> Import File... -> pick this .bas
'    3. Run macro: AddTimingColumns (F5)
'    4. The macro asks which header to classify (default: "OpeningDate")
'       and adds 7 columns to the RIGHT of your existing data:
'           HourLocal, Weekday, IsWeekend,
'           IsHolidayFR, IsHolidayPL, IsHolidayIN,
'           IsOffHours
'    5. Filter / pivot in Excel as needed. Save with Ctrl+S when satisfied.
'
'  READ-ONLY on the source columns. The macro NEVER modifies any existing
'  cell - it only appends new columns to the right of the last data column.
'  It does NOT save the workbook (you save explicitly with Ctrl+S).
'
'  Holiday calendars duplicated from find-offhours-emails.bas. Keep the
'  two in sync if you edit a holiday list - both projects share the rule
'  set but live in separate VBA projects (Outlook vs. Excel).
' ============================================================================

Option Explicit

' -------------------- CONFIG (edit these if needed) --------------------

' Default header (case-insensitive) of the column carrying the date to classify.
' The macro shows this as the InputBox default; you can override per-run.
Private Const DEFAULT_DATE_COLUMN_HEADER As String = "OpeningDate"

' Row containing the headers (most exports have headers on row 1)
Private Const HEADER_ROW As Long = 1

' Working hours - same as find-offhours-emails.bas
Private Const WORK_START_HOUR As Integer = 9        ' 09:00 local
Private Const WORK_END_HOUR As Integer = 18         ' 18:00 local (6 PM)
Private Const TREAT_SATURDAY_AS_OFFHOURS As Boolean = True
Private Const TREAT_SUNDAY_AS_OFFHOURS As Boolean = True

' India team support coverage end time on a Polish holiday (Paris local time).
' Mumbai works 18:00 IST. IST is UTC+5:30 fixed (no DST). Paris is UTC+1
' (CET, winter) or UTC+2 (CEST, summer). So 18:00 IST = 13:30 Paris winter,
' 14:30 Paris summer. The macro auto-detects DST via IsSummerTime().
Private Const INDIA_END_SUMMER_MIN As Long = 14 * 60 + 30   ' 14:30 = 870
Private Const INDIA_END_WINTER_MIN As Long = 13 * 60 + 30   ' 13:30 = 810

' Holiday calendars - keep in sync with find-offhours-emails.bas
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
' -----------------------------------------------------------------------


Public Sub AddTimingColumns()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "No active sheet. Open the workbook first.", vbExclamation, "Add Timing Columns"
        Exit Sub
    End If

    ' Ask which header to classify (pre-filled with the default)
    Dim colName As String
    colName = InputBox("Column header to classify (case-insensitive):", _
                       "Add Timing Columns", DEFAULT_DATE_COLUMN_HEADER)
    If Len(colName) = 0 Then Exit Sub   ' user pressed Cancel
    colName = Trim(colName)

    ' Find the date column on the active sheet
    Dim dateCol As Long
    dateCol = FindHeaderColumn(ws, colName, HEADER_ROW)
    If dateCol = 0 Then
        MsgBox "Could not find a column with header """ & colName & """ on row " & _
               HEADER_ROW & " of sheet """ & ws.Name & """.", vbExclamation, "Add Timing Columns"
        Exit Sub
    End If

    ' Determine data range using the date column's last filled row
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, dateCol).End(xlUp).Row
    If lastRow <= HEADER_ROW Then
        MsgBox "No data rows found below the header in column " & ColumnLetter(dateCol) & ".", _
               vbExclamation, "Add Timing Columns"
        Exit Sub
    End If
    Dim numRows As Long
    numRows = lastRow - HEADER_ROW

    ' Headers we will add
    Dim outHdrs As Variant
    outHdrs = Array("HourLocal", "Weekday", "IsWeekend", _
                    "IsHolidayFR", "IsHolidayPL", "IsHolidayIN", "IsOffHours")

    ' Idempotency guard: refuse if any of our headers already exist on the sheet
    Dim existingLastCol As Long
    existingLastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    Dim col As Long, existingHdr As String, h As Variant
    For col = 1 To existingLastCol
        existingHdr = LCase(Trim(CStr(ws.Cells(HEADER_ROW, col).Value)))
        For Each h In outHdrs
            If existingHdr = LCase(CStr(h)) Then
                MsgBox "A column named """ & h & """ already exists on this sheet. " & _
                       "Delete the existing classification columns before re-running, " & _
                       "or run on a fresh copy of the file.", vbExclamation, "Add Timing Columns"
                Exit Sub
            End If
        Next h
    Next col

    ' First column to write into = rightmost non-empty header column + 1
    Dim firstNewCol As Long
    firstNewCol = existingLastCol + 1

    ' Read source date column into a 2D array
    Dim srcRange As Range
    Set srcRange = ws.Range(ws.Cells(HEADER_ROW + 1, dateCol), ws.Cells(lastRow, dateCol))
    Dim srcData As Variant
    srcData = srcRange.Value
    ' Single-row ranges return a scalar instead of a 2D array - normalise
    If Not IsArray(srcData) Then
        Dim oneVal As Variant
        oneVal = srcData
        ReDim srcData(1 To 1, 1 To 1)
        srcData(1, 1) = oneVal
    End If

    ' Classify
    Dim outData() As Variant
    ReDim outData(1 To numRows, 1 To 7)
    Dim badRows As Long, offHoursCount As Long
    Dim r As Long
    For r = 1 To numRows
        Dim dt As Date, ok As Boolean
        ok = TryParseDate(srcData(r, 1), dt)
        If ok Then
            Dim hh As Integer, isW As Boolean
            Dim hFR As Boolean, hPL As Boolean, hIN As Boolean, isOff As Boolean
            Dim minuteOfDay As Long, outsideOffice As Boolean, indiaEndMin As Long
            hh = Hour(dt)
            isW = IsWeekendDay(dt)
            hFR = IsInDateList(dt, HOLIDAYS_FR_CSV)
            hPL = IsInDateList(dt, HOLIDAYS_PL_CSV)
            hIN = IsInDateList(dt, HOLIDAYS_IN_MUMBAI_CSV)
            minuteOfDay = hh * 60 + Minute(dt)
            outsideOffice = (minuteOfDay < WORK_START_HOUR * 60 Or _
                             minuteOfDay >= WORK_END_HOUR * 60)
            ' IsOffHours rule:
            '   - Saturday/Sunday => off-hours all day
            '   - Outside 09:00-18:00 Paris local => off-hours
            '   - PL holiday during office hours, AND (IN holiday OR past India
            '     coverage end) => off-hours. India covers PL-holiday mornings
            '     until 14:30 Paris in summer / 13:30 in winter.
            ' France and Mumbai holiday flags are tracked but do NOT drive isOff;
            ' Mumbai only matters here as the India-team-also-off check.
            If IsSummerTime(dt) Then
                indiaEndMin = INDIA_END_SUMMER_MIN
            Else
                indiaEndMin = INDIA_END_WINTER_MIN
            End If
            isOff = (isW And TREAT_SATURDAY_AS_OFFHOURS And Weekday(dt, vbMonday) = 6) Or _
                    (isW And TREAT_SUNDAY_AS_OFFHOURS And Weekday(dt, vbMonday) = 7) Or _
                    outsideOffice Or _
                    (hPL And (hIN Or minuteOfDay >= indiaEndMin))
            outData(r, 1) = hh
            outData(r, 2) = WeekdayName(Weekday(dt, vbMonday), True)
            outData(r, 3) = isW
            outData(r, 4) = hFR
            outData(r, 5) = hPL
            outData(r, 6) = hIN
            outData(r, 7) = isOff
            If isOff Then offHoursCount = offHoursCount + 1
        Else
            badRows = badRows + 1
            ' leave row blank (Empty entries write as empty cells)
        End If
    Next r

    ' Write headers (bold)
    Dim hi As Long
    For hi = 0 To UBound(outHdrs)
        ws.Cells(HEADER_ROW, firstNewCol + hi).Value = outHdrs(hi)
    Next hi
    ws.Range(ws.Cells(HEADER_ROW, firstNewCol), _
             ws.Cells(HEADER_ROW, firstNewCol + UBound(outHdrs))).Font.Bold = True

    ' Bulk-write data
    ws.Range(ws.Cells(HEADER_ROW + 1, firstNewCol), _
             ws.Cells(lastRow, firstNewCol + UBound(outHdrs))).Value = outData

    ' Auto-fit just the new columns
    ws.Range(ws.Cells(HEADER_ROW, firstNewCol), _
             ws.Cells(lastRow, firstNewCol + UBound(outHdrs))).Columns.AutoFit

    ' Summary
    Dim msg As String
    msg = "Sheet:        " & ws.Name & vbCrLf & _
          "Date column:  " & colName & " (col " & ColumnLetter(dateCol) & ")" & vbCrLf & _
          "Rows scanned: " & numRows & vbCrLf & _
          "  Off-hours (IsOffHours=TRUE): " & offHoursCount & vbCrLf
    If badRows > 0 Then
        msg = msg & "  Rows with unparseable dates: " & badRows & " (left blank)" & vbCrLf
    End If
    msg = msg & vbCrLf & _
          "Added columns " & ColumnLetter(firstNewCol) & ":" & ColumnLetter(firstNewCol + UBound(outHdrs)) & ":" & vbCrLf & _
          "  HourLocal, Weekday, IsWeekend," & vbCrLf & _
          "  IsHolidayFR, IsHolidayPL, IsHolidayIN," & vbCrLf & _
          "  IsOffHours" & vbCrLf & vbCrLf & _
          "Workbook NOT auto-saved. Press Ctrl+S to save."
    MsgBox msg, vbInformation, "Add Timing Columns"
End Sub


' Find the column index whose header (in HEADER_ROW) case-insensitively
' matches the given name. Returns 0 if not found.
Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerName As String, ByVal headerRow As Long) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    Dim target As String
    target = LCase(Trim(headerName))
    Dim col As Long
    For col = 1 To lastCol
        If LCase(Trim(CStr(ws.Cells(headerRow, col).Value))) = target Then
            FindHeaderColumn = col
            Exit Function
        End If
    Next col
End Function


' Best-effort coercion of a cell value to a Date. Returns True on success
' with the parsed value in dt. Handles native Date cells, numeric serials,
' and parseable text. Returns False for empty / unparseable values.
Private Function TryParseDate(ByVal v As Variant, ByRef dt As Date) As Boolean
    On Error Resume Next
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Len(CStr(v)) = 0 Then Exit Function
    Dim parsed As Date
    parsed = CDate(v)
    If Err.Number = 0 Then
        dt = parsed
        TryParseDate = True
    End If
    Err.Clear
End Function


Private Function IsWeekendDay(ByVal t As Date) As Boolean
    Dim dow As Integer
    dow = Weekday(t, vbMonday)
    IsWeekendDay = (dow = 6 Or dow = 7)
End Function


' Generic membership check against a YYYY-MM-DD CSV list (date only - ignores time-of-day).
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


' Convert a 1-based column index to its Excel letter representation (1->A, 27->AA, ...).
Private Function ColumnLetter(ByVal col As Long) As String
    Dim s As String
    Do While col > 0
        s = Chr(((col - 1) Mod 26) + 65) & s
        col = (col - 1) \ 26
    Loop
    ColumnLetter = s
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
