VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ParticipantDate 
   Caption         =   "Participant and Date"
   ClientHeight    =   2475
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
Public mainWB As Workbook, thisWB As Workbook, participants As Range, theDate As Date, _
        theResult As Boolean

Private Sub Participant_Change()
    
   
    Debug.Print Participant.Value
    
    fromForm (Participant.Value)
    
        
End Sub


Private Sub StatsDate_Change()
    
    theDate = StatsDate.Value
    
    Debug.Print (theDate)
    
End Sub


Sub UserForm_Activate()
    
    'load the participants from data sheet
    
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
    
    Participant.RowSource = participants.Address
    
    'mainWB.Worksheets("Data").Range("b15").Select
    'Set participants = Range("PartIndex")
    


End Sub

Sub fromForm(offIdx)
        Dim partName As String, theDateRange As Range, sheetsToCheck As Variant, badDates As Boolean
        
        partName = participants.Value2(offIdx, 2) & " " & participants.Value2(offIdx, 3)
        
        msg = "work on " & partName & "?"
        
        Debug.Print partName

        response = MsgBox(msg, vbOKCancel)
        
        If response = vbOK Then
            On Error Resume Next
            
            ' Open the file dialog
            With Application.FileDialog(msoFileDialogOpen)
                   .AllowMultiSelect = False
                   .Title = partName
                   .InitialFileName = "C:\Users\Mark\OneDrive\Spring 2016 ILP\Participant Games\" & partName & "\Statistics"
                   ' .InitialFileName = "C:\Users\mark_\OneDrive\Spring 2016 ILP\Participant Games\" & partName & "\Statistics"
                   .Show
            
                   Set thisWB = Workbooks.Open(.SelectedItems(1))
                   
            End With

            Debug.Print thisWB.name
                       
            Set thisWB = Workbooks(partName & " ILP Stats.xlsx")
                            
            ' Set thisWB = ActiveWorkbook
                            
            thisWB.Activate
            
            sheetsToCheck = Array("Assisting Agreements", "Guests", "Registrations")
                        
            ' check for bad dates
            
            For Each sheetname In sheetsToCheck
                
                Debug.Print sheetname
                
                theResult = False
                
                thisWB.Activate
                
                thisWB.Worksheets(sheetname).Activate
                
                Range("B5").Select
                If Not Selection = "" Then
                    If Not Selection.Offset(1, 0).Value = "" Then
                        Range(Selection, Selection.End(xlDown)).Select
                    
                        Set theDateRange = Range(Selection, Selection.End(xlDown))
                        
                        Call checkText(theDateRange, thisWB)
                        
                        Debug.Print "result "; theResult
                        
                        If theResult Then Exit Sub
                        
                    End If
                Else
                    Debug.Print sheetname; " has no data"
                End If
            
            Next sheetname
                   
            response = MsgBox("copy stats?", vbOKCancel)
            
            If response = vbOK Then
                copyStats (offIdx - 1)
                ' thisWB.Close savechanges:=False
            
            Else
                thisWB.Activate
                Exit Sub
            
            End If
        End If

        mainWB.Save
        

End Sub

 Private Sub copyStats(offIdx)
'
' copyStats Macro
'
    
    mainWBName = "CAL ILP Stats 2016-04-15.xlsx"
    
'    offIdx = 10

    ' If mainWB Is Nothing Then Set mainWB = Workbooks(mainWBName)
    
    'Set thisWB = ActiveWorkbook
    
    
    Debug.Print thisWB.name; " index "; offIdx

'   Game
    thisWB.Activate
    
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


Sub checkText(theRange, checkbook)
    
    ' return true if there are bad dates
    
    Dim checksheet As Worksheet
    
    Debug.Print "function starts with this workbook active "; ActiveWorkbook.name
    
    Debug.Print theRange.Parent.Parent.name
    
    ' Set checkbook = Workbooks(theRange.Parent.Parent.name)
    Set checksheet = checkbook.Worksheets(theRange.Parent.name)
    checkbook.Activate
    checksheet.Activate
    
    For Each ddate In theRange
        Debug.Print ddate.Value; ddate.Address; WorksheetFunction.IsText(ddate)

        
        
        If Not WorksheetFunction.IsNumber(ddate) _
            Or WorksheetFunction.IsText(ddate) Then
            
            MsgBox ("text date at " & ActiveSheet.name & " " & ddate.Address)
            checkbook.Activate
            checksheet.Activate
            Range(ddate.Address).Activate
            
            theResult = True
            Exit Sub
            
        ElseIf ddate < Range("ProgramStart") - 21 Or ddate > Worksheets("Schedule").Range("b34") Then
            
            MsgBox ("date out of range at " & ActiveSheet.name & " " & ddate.Address)
            checkbook.Activate
            checksheet.Activate
            Range(ddate.Address).Activate
            
            theResult = True
            Exit Sub
        
        End If
        
    Next ddate
    
    
End Sub


