Attribute VB_Name = "Module1"

' Module: GroupTools
Option Explicit

' === CONFIG ===
Private Const HEADER_NAME As String = "osm_id"  ' header text in row 1
Private Const SUMMARY_ROW_ABOVE As Boolean = True   ' True = summary row above groups; False = below
Private Const GROUP_SINGLETONS As Boolean = False   ' True = also group blocks of 1 row
Private Const GROUP_EMPTY_KEYS As Boolean = False   ' True = also group empty osm_id blocks

' Entry point: group contiguous blocks by osm_id
Public Sub GroupRowsByOSMId()
    Dim ws As Worksheet
    Set ws = ActiveSheet  ' or: Set ws = ThisWorkbook.Worksheets("YourSheetName")

    Dim lastRow As Long, lastCol As Long, keyCol As Long
    If ws Is Nothing Then Exit Sub

    ' Find used range boundaries
    If IsEmpty(ws.UsedRange) Then Exit Sub
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then Exit Sub  ' header only

    ' Locate the osm_id column in header row (case-insensitive)
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

    ' Walk through data rows and group contiguous blocks with same osm_id
    Dim r As Long, startRow As Long
    Dim currKey As Variant, nextKey As Variant

    startRow = 2
    For r = 2 To lastRow
        currKey = ws.Cells(r, keyCol).Value

        If r < lastRow Then
            nextKey = ws.Cells(r + 1, keyCol).Value
        Else
            nextKey = "__END__"
        End If

        ' If block boundary reached (key changes OR end of data)
        If Not KeysEqual(currKey, nextKey) Then
            ' Optionally skip empty key blocks
            If Not GROUP_EMPTY_KEYS Then
                If IsEmpty(currKey) Or Len(Trim$(CStr(currKey))) = 0 Then
                    startRow = r + 1
                    GoTo ContinueLoop
                End If
            End If

            ' Only group if more than 1 row in block (unless GROUP_SINGLETONS=True)
            If GROUP_SINGLETONS Or r > startRow Then
                ws.Rows(startRow & ":" & r).Group
            End If
            startRow = r + 1
        End If
ContinueLoop:
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

' Compares two keys for block continuity; treats both empty as equal.
Private Function KeysEqual(a As Variant, b As Variant) As Boolean
    Dim sa As String, sb As String
    If IsEmpty(a) And IsEmpty(b) Then
        KeysEqual = True
        Exit Function
    End If
    sa = Trim$(CStr(a))
    sb = Trim$(CStr(b))
    KeysEqual = (StrComp(sa, sb, vbTextCompare) = 0)
End Function

