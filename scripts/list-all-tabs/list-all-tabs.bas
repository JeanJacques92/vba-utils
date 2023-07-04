' V2 Selected workbok
Sub ListAllTabs()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim link As Hyperlink
    Dim i As Integer
    
    ' Get the currently selected workbook
    Dim selectedWorkbook As Workbook
    Set selectedWorkbook = ActiveWorkbook
    
    ' Create a new sheet in the selected workbook to output the worksheet names
    Set newWs = selectedWorkbook.Sheets.Add(After:= _
             selectedWorkbook.Sheets(selectedWorkbook.Sheets.Count))
    newWs.Name = "Sheet Index"
    i = 1
    
    ' Loop through each worksheet in the selected workbook
    For Each ws In selectedWorkbook.Worksheets
        ' Output the worksheet name to the new sheet
        newWs.Cells(i, 1).Value = ws.Name
        
        ' Create a hyperlink to the worksheet
        Set link = newWs.Hyperlinks.Add(newWs.Cells(i, 1), "", _
                SubAddress:="'" & ws.Name & "'!A1", _
                ScreenTip:="Go to " & ws.Name, _
                TextToDisplay:=ws.Name)
        
        i = i + 1
    Next ws
    
    ' Autofit the columns in the new sheet
    newWs.Columns.AutoFit
End Sub
