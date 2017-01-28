Attribute VB_Name = "ILsNewMonth"
Sub CreateNewMonth()
'
'
    ActiveWorkbook.Sheets("Introduction Leader Roster").Activate
    
    Call FilterAllILs

    Range("ILInfo[Introduction Leader]").Copy

    ActiveWorkbook.Sheets("Put Results Here").Activate


    Range("Results[[#Headers],[Introduction Leader]]").End(xlDown).Offset(1, 0).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("Results[[#Headers],[Start]]").End(xlDown).Select
    
    
    newmonthDate = WorksheetFunction.EoMonth(Selection.Value, 0) + 1
    
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Value = newmonthDate
    
    Range("Results[[#Headers],[End]]").End(xlDown).Select
    Range(Selection, Selection.End(xlDown)).FillDown
    
    
End Sub


