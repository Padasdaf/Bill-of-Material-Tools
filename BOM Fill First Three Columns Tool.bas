Attribute VB_Name = "Module7"
Sub FillFirstThreeColumns()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("BOM + Item")
    lastRow = ws.UsedRange.Rows.Count

    For i = 4 To lastRow
        ws.Cells(i, 1).Value = �Y�
    ws.Cells(i, 2).Value = (i - 3)
    ws.Cells(i, 3).Value = �P14�
    Next i

End Sub
Sub FillFirstThreeColumnsButton()
    Call FillFirstThreeColumns
End Sub

