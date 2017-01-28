Attribute VB_Name = "FilterAllILsModule"
Public Sub FilterAllILs()
'
'
    Range("ILInfo").Select
    Selection.AutoFilter
    ActiveSheet.ListObjects("ILInfo").Range.AutoFilter Field:=2, Criteria1:= _
        Array("Active"), _
        Operator:=xlFilterValues
    ActiveWorkbook.Worksheets("Introduction Leader Roster").ListObjects("ILInfo"). _
        Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Introduction Leader Roster").ListObjects("ILInfo"). _
        Sort.SortFields.Add Key:=Range("ILInfo[[#All],[Introduction Leader]]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Introduction Leader Roster").ListObjects("ILInfo" _
        ).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Introduction Leader Roster").ListObjects("ILInfo"). _
        Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Introduction Leader Roster").ListObjects("ILInfo"). _
        Sort.SortFields.Add Key:=Range("ILInfo[[#All],[Status]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Introduction Leader Roster").ListObjects("ILInfo" _
        ).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


