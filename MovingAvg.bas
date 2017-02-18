Attribute VB_Name = "MovingAvg"
Sub MovingAvg()
Attribute MovingAvg.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MovingAvg Macro
'

'
    Dim newRange As Range
    
    ActiveWorkbook.Sheets("Put Results Here").Activate
    Range("Results").Select
    'ActiveSheet.ShowAllData
        
        
      
    
    Range("Results[#All]").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:= _
        Range("N2:O3"), CopyToRange:=Range( _
        "ForMoving[[#Headers],[Introduction Leader]:[IDX]]"), Unique:=True
    Range("ForMoving").CurrentRegion.Select
    
    Set newRange = Range("ForMoving[[#Headers],[Introduction Leader]]").CurrentRegion
    
    newRange.Select
    
    ActiveSheet.ListObjects("ForMoving").Resize newRange
    
    ActiveWorkbook.Worksheets("Put Results Here").ListObjects("ForMoving").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Put Results Here").ListObjects("ForMoving").Sort. _
        SortFields.Add Key:=Range("ForMoving[Introduction Leader]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Put Results Here").ListObjects("ForMoving").Sort. _
        SortFields.Add Key:=Range("ForMoving[Start]"), SortOn:=xlSortOnValues, _
        Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Put Results Here").ListObjects("ForMoving"). _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Sheets("Moving Average").Select
    ActiveSheet.PivotTables("MovingAvg").PivotCache.Refresh
    Sheets("Put Results Here").Select
    Selection.AutoFilter
    ActiveSheet.ListObjects("Results").Range.AutoFilter Field:=3, Criteria1:= _
        xlFilterLastWeek, Operator:=xlFilterDynamic
        
End Sub
