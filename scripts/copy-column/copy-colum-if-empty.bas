Excel Macro to replace a colum value from other if it is empty:
Sub CopyFormulaResults()
    Dim lastRow As Long
    Dim sourceColumn As Integer
    Dim targetColumn As Integer
    Dim sourceRange As Range
    Dim destinationRange As Range
    
    ' Set the column numbers of the source and target columns
    sourceColumn = 24 ' Change this to the column number of your source column
    targetColumn = 17 ' Change this to the column number of your target column
    
    ' Get the last row in the active sheet
    lastRow = Cells(Rows.Count, sourceColumn).End(xlUp).Row
    
    ' Set the range of the source column
    Set sourceRange = Range(Cells(2, sourceColumn), Cells(lastRow, sourceColumn))
    
    ' Set the range of the target column
    Set destinationRange = Range(Cells(2, targetColumn), Cells(lastRow, targetColumn))
    
    ' Loop through each cell in the source range
    
    
    For i = 2 To 2828
     'MsgBox Cells(i, sourceColumn).Text
     If Cells(i, sourceColumn).Value <> "" And Cells(i, targetColumn).Value = "" Then
        Cells(i, targetColumn) = Cells(i, sourceColumn).Text
     End If
    Next
    ' Clear the clipboard
    Application.CutCopyMode = False
End Sub
