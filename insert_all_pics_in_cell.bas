Attribute VB_Name = "Module1"
Sub PutAllPicsInCells()
    Dim Pic As Shape
    For Each Pic In ActiveSheet.Shapes
        If Pic.Type = msoPicture Then
        Pic.Select
        Pic.PlacePictureInCell
        End If
    Next Pic
End Sub
