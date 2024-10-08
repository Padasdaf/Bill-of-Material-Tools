Attribute VB_Name = "Module6"
Sub UndoDescriptionModify()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim description As String
    Dim semiColonPos As Integer
    
    Set ws = ThisWorkbook.Sheets("BOM + Item")
    lastRow = ws.UsedRange.Rows.Count
    ws.Range("Q4:Q� & lastRow).ClearContents

    For i = 4 To lastRow
        description = ws.Cells(i, 10).Value
        semiColonPos = InStr(description, ";")
        
        If semiColonPos > 0 Then
            ws.Cells(i, 10).Value = Left(description, semiColonPos - 1)
        End If
    Next i

End Sub
Sub UndoDescriptionModifyButton()
    Call UndoDescriptionModify
End Sub

