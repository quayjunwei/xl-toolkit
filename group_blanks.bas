' Module: GroupTools
Option Explicit

' === CONFIG ===
Private Const HEADER_NAME As String = "site_id"     ' header text in row 1
Private Const SUMMARY_ROW_ABOVE As Boolean = True   ' True = summary row above groups; False = below

' Entry point: group contiguous blocks by site_id
' Groups rows from the first non-blank site_id until the next non-blank site_id.
' The first row of each group (where site_id is filled) acts as the summary row.
' Subsequent blank site_id rows (parcel rows) are grouped beneath it.
Public Sub GroupRowsBySiteId()
    Dim ws As Worksheet
    Set ws = ActiveSheet  ' or: Set ws = ThisWorkbook.Worksheets("YourSheetName")

    Dim lastRow As Long, keyCol As Long
    If ws Is Nothing Then Exit Sub

    ' Find used range boundaries
    If IsEmpty(ws.UsedRange) Then Exit Sub
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub  ' header only

    ' Locate the site_id column in header row (case-insensitive)
    keyCol = FindHeaderColumn(ws, HEADER_NAME)
    If keyCol = 0 Then
        MsgBox "Header '" & HEADER_NAME & "' not found in row 1.", vbExclamation
        Exit Sub
    End If

    ' Performance toggles
    Dim oldScrUp As Boolean, oldEvt As Boolean
    Dim oldCalc As XlCalculation
    oldScrUp = Application.ScreenUpdating
    oldEvt = Application.EnableEvents
    oldCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    ' Clear any existing outline and set summary row position
    ws.Cells.ClearOutline
    If SUMMARY_ROW_ABOVE Then
        ws.Outline.SummaryRow = xlAbove
    Else
        ws.Outline.SummaryRow = xlBelow
    End If

    ' Walk through data rows
    ' A new group starts whenever site_id is non-blank (except the very first data row)
    ' Group the parcel rows (startRow+1 to r-1) under each site row
    Dim r As Long, startRow As Long
    Dim cellVal As Variant
    Dim isNewGroup As Boolean

    startRow = 2

    For r = 2 To lastRow
        cellVal = ws.Cells(r, keyCol).Value
        isNewGroup = False

        ' New group starts at any non-blank site_id after the first data row
        If r > 2 Then
            If Not IsEmpty(cellVal) And Len(Trim$(CStr(cellVal))) > 0 Then
                isNewGroup = True
            End If
        End If

        ' Close off the previous group before starting a new one
        If isNewGroup Then
            If r - 1 > startRow Then
                ws.Rows((startRow + 1) & ":" & (r - 1)).Group
            End If
            startRow = r
        End If

        ' Close off the final group at end of data
        If r = lastRow Then
            If lastRow > startRow Then
                ws.Rows((startRow + 1) & ":" & lastRow).Group
            End If
        End If
    Next r

CleanExit:
    ' Restore application state
    Application.ScreenUpdating = oldScrUp
    Application.EnableEvents = oldEvt
    Application.Calculation = oldCalc
    Exit Sub

CleanFail:
    ' Still restore app state on error
    Resume CleanExit
End Sub

' --- Helpers ---

' Finds the column index where row 1 equals the given header name (case-insensitive exact match).
Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(CStr(ws.Cells(1, c).Value), headerName, vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
    FindHeaderColumn = 0
End Function
