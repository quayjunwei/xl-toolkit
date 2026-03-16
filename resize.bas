Sub FixImageRowHeights()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim cell As Range
    For Each cell In ws.Columns(12).SpecialCells(xlCellTypeConstants)
        cell.EntireRow.RowHeight = 192.5  ' adjust as needed
    Next cell
    
    MsgBox "Done!"
End Sub