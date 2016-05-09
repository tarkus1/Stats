Attribute VB_Name = "FilterILsModule"
Public Sub FilterILs()
Attribute FilterILs.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FilterILs Macro
'

'
    Selection.AutoFilter
    ActiveSheet.ListObjects("ILInfo").Range.AutoFilter Field:=2, Criteria1:= _
        Array("Fulfilled Leading to 100! as Program Leader", _
        "Led to 100! Undeclared or Unrealilzed Program Leader", "Not Yet Led to 100!"), _
        Operator:=xlFilterValues
    ActiveWorkbook.Worksheets("Introduction Leader Info").ListObjects("ILInfo"). _
        Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Introduction Leader Info").ListObjects("ILInfo"). _
        Sort.SortFields.Add Key:=Range("ILInfo[[#All],[Introduction Leader]]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Introduction Leader Info").ListObjects("ILInfo" _
        ).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Introduction Leader Info").ListObjects("ILInfo"). _
        Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Introduction Leader Info").ListObjects("ILInfo"). _
        Sort.SortFields.Add Key:=Range("ILInfo[[#All],[Status]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Introduction Leader Info").ListObjects("ILInfo" _
        ).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
