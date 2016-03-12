Attribute VB_Name = "ILPStatsSprint2016"
Public mainWB As Workbook, mainWBName As String

Public partNames As Variant, offIdx As Integer

Public partWBs As Workbooks
Public thisWB As Workbook

    Type particInfo
        name As String
        index As Integer
    End Type


' create the name list and index


' create the name list and run stats

Sub nameList()
    
    Dim keepgoing As Boolean, response As Integer, fileName As String
    

    Dim participants(0 To 11) As particInfo
    
    participants(0).name = "Curtis Arnold"
    participants(1).name = "Judy Bartholomeusz"
    participants(2).name = "Chelsea Bayly-Williams"
    participants(3).name = "Sheila Braun"
    participants(4).name = "Erick Corzo"
    participants(5).name = "Sarah Fletcher"
    participants(6).name = "Fahreen Lalani"
    participants(7).name = "David Lyon"
    participants(8).name = "Donna Mathezing"
    participants(9).name = "Chris Munstermann"
    participants(10).name = "Julian Ruth"
    participants(11).name = "Audrey Wilkins"
    
    participants(0).index = 0
    participants(1).index = 1
    participants(2).index = 2
    participants(3).index = 3
    participants(4).index = 4
    participants(5).index = 5
    participants(6).index = 6
    participants(7).index = 7
    participants(8).index = 8
    participants(9).index = 9
    participants(10).index = 10
    participants(11).index = 11
    
    
    mainWBName = "CAL ILP Stats 2016-03-11.xlsx"
    Set mainWB = Workbooks(mainWBName)
    


    offIdx = 0
    
    Do While offIdx < 12
        Debug.Print participants(offIdx).name
        offIdx = offIdx + 1
        
    Loop

   offIdx = 0
        
    Do While offIdx < 12
        msg = "work on " & participants(offIdx).name & "?"
       response = MsgBox(msg, vbOKCancel)
        If response = vbOK Then
            On Error Resume Next
            
            fileName = "C:\Users\Mark\OneDrive\Spring 2016 ILP\Participant Games\" & participants(offIdx).name & _
                            "\Statistics\" & participants(offIdx).name & " ILP Stats.xlsx"
            
            'fileName = "C:\Users\mark_\OneDrive\Participant Games\" & participants(offIdx).name & _
                            "\Statistics\ILP Stats " & participants(offIdx).name & ".xlsx"
            
            Debug.Print fileName
            
            Workbooks.Open fileName
                            
            Set thisWB = Workbooks(participants(offIdx).name & " ILP Stats.xlsx")
            
            thisWB.Activate
            
            response = MsgBox("copy stats?", vbOKCancel)
            
            If response = vbOK Then
                copyStats (offIdx)
                thisWB.Close savechanges:=False
            
            Else
                thisWB.Activate
                Exit Sub
            
            End If
        End If
                
        
        offIdx = offIdx + 1
        
    Loop
         
    End Sub
'Sub copyStats()

 Sub copyStats(offIdx)
'
' copyStats Macro
'
    
    mainWBName = "CAL ILP Stats 2016-03-11.xlsx"
    
'    offIdx = 10

    If mainWB Is Nothing Then Set mainWB = Workbooks(mainWBName)
    
    Set thisWB = ActiveWorkbook
    
    
    Debug.Print thisWB.name; " index "; offIdx

'   Game

    thisWB.Worksheets("Statistician").Activate
    
    Range("A15:gf15").Select
    Selection.Copy
    
    mainWB.Worksheets("Data").Activate

    Range("G15").Offset(offIdx, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'   assignments

    thisWB.Worksheets("Statistician").Activate
    
    Range("b7:be7").Select
    Selection.Copy
    
    mainWB.Worksheets("Assignments").Activate

    Range("g5").Offset(offIdx, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'weekly measures
    
    thisWB.Worksheets("Statistician").Activate
    
    Range("a23:bh23").Select
    Selection.Copy
    
    mainWB.Worksheets("WeeklyMeasures").Activate

    Range("g7").Offset(offIdx, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    mainWB.Save
    
End Sub

Sub listWB()
    For Each thisWB In Workbooks
        Debug.Print thisWB.name
    Next thisWB
End Sub


Sub closeWB()
'
' closeWB Macro
'

'
    thisWB.Close savechanges:=False
    
End Sub


