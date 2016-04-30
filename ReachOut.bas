Attribute VB_Name = "ReachOut"
Public participants As Range, mainWB As Workbook, partName As String, i As Integer, _
    thisWB As Workbook, touchName As Variant
Sub ReachOut()

    For Each wb In Application.Workbooks
        Debug.Print wb.name
        If Left(wb.name, 7) = "CAL ILP" Then Set mainWB = wb
    Next wb
    
    If mainWB Is Nothing Then Exit Sub
    
    mainWB.Sheets("Data").Activate
    
       
    Range("C15").Select
    
    mainWB.Sheets("Data").Range("C15", Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    
    Set participants = Selection
    
    For i = 1 To participants.Rows.Count
    
        partName = participants.Value2(i, 2) & " " & participants.Value2(i, 3)
       
        ' Debug.Print partName
        theFile = "C:\Users\Mark\OneDrive\Spring 2016 ILP\Participant Games\" & partName & _
                   "\Statistics\" & partName & " ILP Stats.xlsx"
        ' Debug.Print theFile

       
        Set thisWB = Application.Workbooks.Open(theFile)
    
        thisWB.Worksheets("Reach Out & Touch").Activate
        
        touchName = Application.WorksheetFunction.CountA(Range("B5:B104"))
        
        Debug.Print partName; " count in reach out "; touchName
        
        thisWB.Close savechanges:=False
        
        
    Next i

End Sub

