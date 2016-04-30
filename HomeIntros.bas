Attribute VB_Name = "HomeIntros"
Public participants As Range, mainWB As Workbook, partName As String, i As Integer, _
    thisWB As Workbook, HomeIntros As Range
Sub HomeIntroductions()

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
    
        ' partName = participants.Value2(i, 2) & " " & participants.Value2(i, 3)
       
        partName = "Chris Munstermann"
        
        ' Debug.Print partName
        theFile = "C:\Users\Mark\OneDrive\Spring 2016 ILP\Participant Games\" & partName & _
                   "\Statistics\" & partName & " ILP Stats.xlsx"
        ' Debug.Print theFile

       
        Set thisWB = Application.Workbooks.Open(theFile)
    
        thisWB.Worksheets("Home Intros").Activate
        
        Range("b6").Select
        
        Range(Selection, Selection.End(xlDown)).Select
        
        Debug.Print homintros.Value2
        
        thisWB.Close savechanges:=False
        
        
    Next i

End Sub



