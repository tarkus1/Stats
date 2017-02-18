Attribute VB_Name = "ReachOut"

Sub ReachOut()
    Dim participants As Range, mainWB As Workbook, partName As String, i As Integer, _
    thisWB As Workbook, tName As Variant, tWB As Workbook, theFile As String
    
    For Each wb In Application.Workbooks
        Debug.Print wb.name
        If Left(wb.name, 7) = "CAL ILP" Then Set mainWB = wb
        If Left(wb.name, 5) = "Suppl" Then Set tWB = wb
    Next wb
    
    If mainWB Is Nothing Or tWB Is Nothing Then Exit Sub
    
    mainWB.Sheets("Data").Activate
    
       
    Range("C15").Select
    
    mainWB.Sheets("Data").Range("C15", Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    
    Set participants = Selection
    
    For i = 1 To participants.Rows.count
    
        partName = participants.Value2(i, 2) & " " & participants.Value2(i, 3)
       
        ' Debug.Print partName
        theFile = "C:\Users\mark_\OneDrive\Fall 2016 ILP\Participant Games\" & partName & _
                   "\Statistics\" & partName & " ILP Stats.xlsx"
        Debug.Print theFile

       
        Set thisWB = Application.Workbooks.Open(theFile)
    
        thisWB.Worksheets("Reach Out & Touch").Activate
        
        tName = Application.WorksheetFunction.CountA(Range("c6:c105"))
        
        tWB.Worksheets("Reach out").Activate
        ActiveSheet.Cells(i + 1, 1).Select
        ActiveSheet.Cells(i + 1, 1).Value = partName
        ActiveSheet.Cells(i + 1, 2).Value = tName
        
        thisWB.Close savechanges:=False
        
        
    Next i

End Sub

