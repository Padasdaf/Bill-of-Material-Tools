Attribute VB_Name = "Module8"
Sub DOrM()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("BOM + Item")
    lastRow = ws.UsedRange.Rows.Count

    For i = 4 To lastRow
    If ws.Cells(i, 7).Value = “PrimarySpec” Then
        ws.Cells(i, 6).Value = “D”
    ElseIf Left(ws.Cells(i, 7).Value, 4) = “ASTM” Then
        ws.Cells(i, 6).Value = “M”
    ElseIf Left(ws.Cells(i, 7).Value, 4) = “B50A” Then
        ws.Cells(i, 6).Value = “M”
    End If
    Next i

End Sub
Sub DOrMButton()
    CallDOrM
End Sub

