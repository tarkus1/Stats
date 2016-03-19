VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ParticipantDate 
   Caption         =   "Participant and Date"
   ClientHeight    =   2480
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3795
   OleObjectBlob   =   "ParticipantDate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ParticipantDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public mainWB As Workbook, thisWB As Workbook, participants As Range

Private Sub Participant_Change()
    Debug.Print Participant.Value
    
    fromForm (Participant.Value)
    
        
End Sub


Sub UserForm_Activate()
    
    'load the participants from data sheet
    
    Set mainWB = ActiveWorkbook
    
    'mainWB.Worksheets("Data").Range("b15").Select
    Set participants = Range("PartIndex")
    


End Sub

Sub fromForm(offIdx)
        Dim partName As String
        
        partName = participants.Value2(offIdx, 2) & " " & participants.Value2(offIdx, 3)
        
        msg = "work on " & partName & "?"
        
        Debug.Print partName

        response = MsgBox(msg, vbOKCancel)
        
        If response = vbOK Then
            On Error Resume Next
            
            fileName = "C:\Users\Mark\OneDrive\Spring 2016 ILP\Participant Games\" & partName & _
                            "\Statistics\" & partName & " ILP Stats.xlsx"
             
            ' fileName = "C:\Users\mark_\OneDrive\Spring 2016 ILP\Participant Games\" & partName & _
                            "\Statistics\" & partName & " ILP Stats.xlsx"
            
           
            'fileName = "C:\Users\mark_\OneDrive\Participant Games\" & participants(offIdx).name & _
                            "\Statistics\ILP Stats " & participants(offIdx).name & ".xlsx"
            
            Debug.Print fileName
            
            Workbooks.Open fileName
                            
            Set thisWB = Workbooks(participants(offIdx).name & " ILP Stats.xlsx")
            
            thisWB.Activate
            
            response = MsgBox("copy stats?", vbOKCancel)
            
            If response = vbOK Then
                copyStats (offIdx - 1)
                thisWB.Close savechanges:=False
            
            Else
                thisWB.Activate
                Exit Sub
            
            End If
        End If




End Sub

 Private Sub copyStats(offIdx)
'
' copyStats Macro
'
    
    mainWBName = "CAL ILP Stats 2016-03-18.xlsx"
    
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
    
    ' mainWB.Save
    
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


