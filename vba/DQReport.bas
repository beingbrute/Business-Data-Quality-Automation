Attribute VB_Name = "DQReport"
' Business Data Quality Automation: Excel VBA post-processing
' Run ProcessLatestReport after validate.py has written a report.
' 1. Finds the newest dq_report_*.xlsx under reports\ (including dated subfolders)
' 2. Adds a summary block to the Summary sheet and links each check to its exception sheet
' 3. Exports the Summary sheet as a one-page PDF next to the report
' 4. Appends one row per run to logs\dq_tracker.xlsx (skips reports already logged)
' Keep this module in vba\DQ_Macros.xlsm inside the repository folder.

Option Explicit

Private Const TRACKER_FILE As String = "dq_tracker.xlsx"
Private Const BLOCK_ROWS As Long = 5

Public Sub ProcessLatestReport()
    Dim repoDir As String, reportPath As String
    repoDir = Left$(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\") - 1)
    reportPath = FindLatestReport(repoDir & "\reports")
    If reportPath = "" Then
        MsgBox "No dq_report_*.xlsx found under " & repoDir & "\reports", vbExclamation
        Exit Sub
    End If
    ProcessReport reportPath, repoDir
End Sub

Private Function FindLatestReport(ByVal folderPath As String) As String
    Dim fso As Object, best As String, bestTime As Date
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Function
    SearchFolder fso.GetFolder(folderPath), best, bestTime
    FindLatestReport = best
End Function

Private Sub SearchFolder(ByVal fld As Object, ByRef best As String, ByRef bestTime As Date)
    Dim f As Object, sf As Object
    For Each f In fld.Files
        If LCase$(f.Name) Like "dq_report_*.xlsx" Then
            If best = "" Or f.DateLastModified > bestTime Then
                best = f.Path
                bestTime = f.DateLastModified
            End If
        End If
    Next f
    For Each sf In fld.SubFolders
        SearchFolder sf, best, bestTime
    Next sf
End Sub

Private Sub ProcessReport(ByVal reportPath As String, ByVal repoDir As String)
    Dim wb As Workbook, ws As Worksheet
    Dim reportName As String, pdfPath As String
    Dim checkCol As Long, statusCol As Long, failedCol As Long
    Dim lastRow As Long, r As Long, nChecks As Long
    Dim nFail As Long, nWarn As Long, nPass As Long, rowsFailed As Long
    Dim sheetName As String

    On Error GoTo Failed
    Application.ScreenUpdating = False
    reportName = Mid$(reportPath, InStrRev(reportPath, "\") + 1)
    Set wb = Workbooks.Open(reportPath)
    Set ws = wb.Worksheets("Summary")

    ' Header must still be in row 1; if not, this report was already processed
    If LCase$(CStr(ws.Range("A1").Value)) <> "table" Then
        wb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox reportName & " has already been processed.", vbInformation
        Exit Sub
    End If

    checkCol = HeaderColumn(ws, "check")
    statusCol = HeaderColumn(ws, "status")
    failedCol = HeaderColumn(ws, "rows_failed")
    If checkCol = 0 Or statusCol = 0 Or failedCol = 0 Then
        Err.Raise vbObjectError + 513, , "Summary sheet is missing the check, status or rows_failed column."
    End If

    lastRow = ws.Cells(ws.Rows.Count, checkCol).End(xlUp).Row
    nChecks = lastRow - 1
    For r = 2 To lastRow
        Select Case CStr(ws.Cells(r, statusCol).Value)
            Case "FAIL": nFail = nFail + 1
            Case "WARN": nWarn = nWarn + 1
            Case "PASS": nPass = nPass + 1
        End Select
        rowsFailed = rowsFailed + CLng(Val(ws.Cells(r, failedCol).Value))
        sheetName = Left$(CStr(ws.Cells(r, checkCol).Value), 31)
        If SheetExists(wb, sheetName) Then
            ws.Hyperlinks.Add Anchor:=ws.Cells(r, checkCol), Address:="", _
                SubAddress:="'" & sheetName & "'!A1", TextToDisplay:=CStr(ws.Cells(r, checkCol).Value)
        End If
    Next r

    ' Summary block above the table
    ws.Rows("1:" & BLOCK_ROWS).Insert
    With ws
        .Range("A1").Value = "Data quality report"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "Report file":    .Range("B2").Value = reportName
        .Range("A3").Value = "Processed on":   .Range("B3").Value = Now
        .Range("B3").NumberFormat = "dd-mmm-yyyy hh:mm"
        .Range("A4").Value = "Checks run":     .Range("B4").Value = nChecks
        .Range("D2").Value = "FAIL":           .Range("E2").Value = nFail
        .Range("D3").Value = "WARN":           .Range("E3").Value = nWarn
        .Range("D4").Value = "PASS":           .Range("E4").Value = nPass
        .Range("G2").Value = "Rows failed":    .Range("H2").Value = rowsFailed
        .Range("A2:A4,D2:D4,G2").Font.Bold = True
        .Range("D2:E2").Interior.Color = RGB(255, 199, 206)
        .Range("D3:E3").Interior.Color = RGB(255, 235, 156)
        .Range("D4:E4").Interior.Color = RGB(198, 239, 206)
    End With

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A" & (BLOCK_ROWS + 2)).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select

    ' One-page landscape PDF of the Summary sheet
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    pdfPath = Left$(reportPath, Len(reportPath) - 5) & "_summary.pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath

    wb.Close SaveChanges:=True
    AppendToTracker repoDir, reportName, nChecks, nFail, nWarn, nPass, rowsFailed
    Application.ScreenUpdating = True

    MsgBox "Processed " & reportName & vbCrLf & _
           "FAIL " & nFail & "  |  WARN " & nWarn & "  |  PASS " & nPass & "  |  rows failed " & rowsFailed & vbCrLf & _
           "PDF: " & pdfPath, vbInformation
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Could not process the report: " & Err.Description, vbCritical
End Sub

Private Sub AppendToTracker(ByVal repoDir As String, ByVal reportName As String, _
                            ByVal nChecks As Long, ByVal nFail As Long, ByVal nWarn As Long, _
                            ByVal nPass As Long, ByVal rowsFailed As Long)
    Dim logDir As String, trackerPath As String
    Dim twb As Workbook, tws As Worksheet, r As Long

    logDir = repoDir & "\logs"
    If Dir(logDir, vbDirectory) = "" Then MkDir logDir
    trackerPath = logDir & "\" & TRACKER_FILE

    If Dir(trackerPath) = "" Then
        Set twb = Workbooks.Add(xlWBATWorksheet)
        Set tws = twb.Worksheets(1)
        tws.Name = "Runs"
        tws.Range("A1:H1").Value = Array("processed_on", "report_file", "checks", "fail", "warn", "pass", "rows_failed", "month")
        tws.Range("A1:H1").Font.Bold = True
        twb.SaveAs Filename:=trackerPath, FileFormat:=xlOpenXMLWorkbook
    Else
        Set twb = Workbooks.Open(trackerPath)
        Set tws = twb.Worksheets("Runs")
    End If

    If Not IsError(Application.Match(reportName, tws.Columns(2), 0)) Then
        twb.Close SaveChanges:=False
        Exit Sub
    End If

    r = tws.Cells(tws.Rows.Count, 1).End(xlUp).Row + 1
    tws.Cells(r, 1).Value = Now
    tws.Cells(r, 1).NumberFormat = "dd-mmm-yyyy hh:mm"
    tws.Cells(r, 2).Value = reportName
    tws.Cells(r, 3).Value = nChecks
    tws.Cells(r, 4).Value = nFail
    tws.Cells(r, 5).Value = nWarn
    tws.Cells(r, 6).Value = nPass
    tws.Cells(r, 7).Value = rowsFailed
    tws.Cells(r, 8).Value = Format$(Now, "yyyy-mm")
    tws.Columns("A:H").AutoFit
    twb.Close SaveChanges:=True
End Sub

Private Function HeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If LCase$(Trim$(CStr(ws.Cells(1, c).Value))) = headerName Then
            HeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim sh As Object
    For Each sh In wb.Worksheets
        If sh.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next sh
End Function
