Attribute VB_Name = "NewWeek"
Sub NewWeek()
Attribute NewWeek.VB_ProcData.VB_Invoke_Func = " \n14"
'
' NewWeek Macro
'

'
    Sheets("Introduction Leader Info").Activate

    Range("ILInfo[Introduction Leader]").Copy

    Sheets("Put Results Here").Activate


    Range("Results[[#Headers],[Introduction Leader]]").End(xlDown).Offset(1, 0).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("Results[[#Headers],[Start]]").End(xlDown).Select
    
    
    newWeekDate = Selection.Value + 7
    
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Value = newWeekDate
    
    
    
    
End Sub
