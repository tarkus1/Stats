Attribute VB_Name = "ILPStats"
Public mainWB As Workbook

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
    
    participants(0).name = "Arjay Ratcliffe"
    participants(1).name = "Arlen Adolphus"
    participants(2).name = "Courtney Lyall"
    participants(3).name = "Erin Hlynsky"
    participants(4).name = "Joy Stevenson"
    participants(5).name = "Julia Pho"
    participants(6).name = "Kim Pearson"
    participants(7).name = "Nicholas Novello"
    participants(8).name = "Raelene Izquierdo"
    participants(9).name = "Richard Wensing"
    participants(10).name = "Rummy Rendina"
    participants(11).name = "Shawn Johnson"
    
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
    

    Set mainWB = Workbooks("Calgary ILP 15-2 Classroom Workbook Week 24.xlsx")
    


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
            'fileName = "C:\Users\Mark\OneDrive\Participant Games\" & participants(offIdx).name & _
                            "\Statistics\ILP Stats " & participants(offIdx).name & ".xlsx"
            
            fileName = "C:\Users\mark_\OneDrive\Participant Games\" & participants(offIdx).name & _
                            "\Statistics\ILP Stats " & participants(offIdx).name & ".xlsx"
            
            Debug.Print fileName
            
            Workbooks.Open fileName
                            
            Set thisWB = Workbooks("ILP Stats " & participants(offIdx).name & ".xlsx")
            
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
    If mainWB Is Nothing Then Set mainWB = Workbooks("Calgary ILP 15-2 Classroom Workbook Week 24.xlsx")
    
    Set thisWB = ActiveWorkbook
    
    
    Debug.Print thisWB.name; " index "; offIdx

'   Game

    thisWB.Worksheets("Statistician").Activate
    
    Range("A15:HJ15").Select
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

    Range("g7").Offset(offIdx, 0).Select
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
Attribute closeWB.VB_ProcData.VB_Invoke_Func = " \n14"
'
' closeWB Macro
'

'
    thisWB.Close savechanges:=False
    
End Sub
