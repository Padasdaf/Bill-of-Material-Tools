Attribute VB_Name = "Module2"
Sub ManufacturingBOMTransferData()
    Dim lastRowSource As Long
    Dim sourceRange As Range
    Dim destinationRange As Range
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim wbSource As Workbook
    Dim wbDestination As Workbook
    Dim i As Long
    Dim sourcePath As String
    Dim destinationPath As String
    SourcePath = "C:\Users\212785994\Desktop\BOM transfer\Manufacturingï?æ?BOM.xls"
    destinationPath = "C:\Users\212785994\Desktop\BOM transfer\GEHZ BOM€£¡Œ.xlsm"
    Set wbSource = Workbooks.Open(sourcePath)
    Set wsSource = wbSource.Sheets("BOM")
    Set wbDestination = Workbooks.Open(destinationPath)
    Set wsDestination = wbDestination.Sheets("BOM + Item")
    lastRowSource = 0
    Dim col As Integer
    For col = 1 To 17
        Dim lastRowInCol As Long
        lastRowInCol = wsSource.Cells(wsSource.Rows.Count, col).End(xlUp).Row
        If lastRowInCol > lastRowSource Then
            lastRowSource = lastRowInCol
        End If
    Next col
    Set sourceRange = wsSource.Range("A4:A" & lastRowSource)
    Set destinationRange = wsDestination.Range("E4")
    sourceRange.Copy Destination:=destinationRange
    For i = 4 To lastRowSource
        If Not IsEmpty(wsSource.Cells(i, 4).Value) Then
            wsDestination.Cells(i, 8).Value = wsSource.Cells(i, 4).Value
        ElseIf Not IsEmpty(wsSource.Cells(i, 12).Value) Then
            wsDestination.Cells(i, 8).Value = wsSource.Cells(i, 12).Value
        Else
            wsDestination.Cells(i, 8).Value = wsSource.Cells(i, 11).Value
        End If
    Next i
    For i = 4 To lastRowSource
        If Not IsEmpty(wsSource.Cells(i, 2).Value) Then
            wsDestination.Cells(i, 7).Value = wsSource.Cells(i, 2).Value
        ElseIf Not IsEmpty(wsSource.Cells(i, 10).Value) Then
            wsDestination.Cells(i, 7).Value = wsSource.Cells(i, 10).Value
        ElseIf Not IsEmpty(wsSource.Cells(i, 11).Value) Then
            wsDestination.Cells(i, 7).Value = wsSource.Cells(i, 11).Value
        Else
            wsDestination.Cells(i, 7).Value = wsSource.Cells(i, 12).Value
        End If
    Next i
    Set sourceRange = wsSource.Range("F4:F" & lastRowSource)
    Set destinationRange = wsDestination.Range("M4")
    sourceRange.Copy Destination:=destinationRange
    Set sourceRange = wsSource.Range("G4:G" & lastRowSource)
    Set destinationRange = wsDestination.Range("L4")
    sourceRange.Copy Destination:=destinationRange
    For i = 4 To lastRowSource
        If Not IsEmpty(wsSource.Cells(i, 14).Value) And wsSource.Cells(i, 14).Value <> "Released" Then
            wsDestination.Cells(i, 10).Value = wsSource.Cells(i, 14).Value
        ElseIf Not IsEmpty(wsSource.Cells(i, 12).Value) Then
            wsDestination.Cells(i, 10).Value = wsSource.Cells(i, 12).Value
        Else
            wsDestination.Cells(i, 10).Value = wsSource.Cells(i, 11).Value
        End If
    Next i
End Sub
Sub ManufacturingButton()
    Call ManufacturingBOMTransferData
End Sub
