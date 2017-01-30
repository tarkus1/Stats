Attribute VB_Name = "NewWeek"
Sub CreateNewWeek()
'
' NewWeek Macro
'

'
    ActiveWorkbook.Sheets("Introduction Leader Info").Activate
    
    Call FilterILs

    Range("ILInfo[Introduction Leader]").Copy

    ActiveWorkbook.Sheets("Put Results Here").Activate


    Range("Results[[#Headers],[Introduction Leader]]").End(xlDown).Offset(1, 0).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("Results[[#Headers],[Start]]").End(xlDown).Select
    
    
    newWeekDate = Selection.Value + 7
    
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Value = newWeekDate
    
    Range("Results[[#Headers],[End]]").End(xlDown).Select
    Range(Selection, Selection.End(xlDown)).FillDown
    
    ActiveSheet.ListObjects("Results").Range.AutoFilter Field:=2, Criteria1:= _
        xlFilterLastWeek, Operator:=xlFilterDynamic
    
End Sub
