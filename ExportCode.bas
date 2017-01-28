Attribute VB_Name = "ExportCode"
Sub xPort()
     
     ' reference to extensibility library
     
    Dim objMyProj As VBProject
    Dim objVBComp As VBComponent
     
    Set objMyProj = Application.VBE.ActiveVBProject
     
    For Each objVBComp In objMyProj.VBComponents
        If objVBComp.Type = vbext_ct_StdModule Then
            Debug.Print "module "; objVBComp.Name
            
            ' objVBComp.Export "C:\Users\Mark\OneDrive\MakeWaves\Landmark\Introduction Leader\Stats" & objVBComp.name & ".bas"
            objVBComp.Export "C:\Users\Mark\OneDrive\MakeWaves\Landmark\Introduction Leader\Stats\" & objVBComp.Name & ".bas"
        End If
    Next
     
End Sub


Private Sub ListModulesAndSubs()
     Dim wb As Workbook
     Dim vbComp As VBComponent
     Dim wsList As Worksheet
     Dim i As Long
     Dim strCodeLine As String
     Dim strProcBodyLine As String
     Dim strModule As String
     
     Set wb = ActiveWorkbook
     'Alternative code to previous line if workbook _
      being processed is not ThisWorkbook containing the code
     'Set wb = Workbooks("My Other Workbook.xlsm")
     
     Set wsList = wb.Sheets("Sheet3")    'Edit "Sheet3" to required sheet for output.
     wsList.Range("A1") = "Module"
     wsList.Range("B1") = "Sub Routine"
     wsList.Range("A1:B1").Font.Bold = True
     
     For Each vbComp In wb.VBProject.VBComponents
         With vbComp.CodeModule
             For i = .CountOfDeclarationLines + 1 To .CountOfLines
                 If Trim(.Lines(i, 1)) <> "" Then
                     If strProcBodyLine <> .Lines _
                             (.ProcBodyLine(.ProcOfLine(i, _
                             vbext_pk_Proc), vbext_pk_Proc), 1) Then
                         
                         strModule = vbComp.CodeModule
                         strCodeLine = .Lines(i, 1)
                         
                         With wsList
                             .Cells(.Rows.count, "A").End(xlUp).Offset(1, 0) _
                                 = strModule
                             .Cells(.Rows.count, "B").End(xlUp).Offset(1, 0) _
                                 = strCodeLine
                         End With
                         
                         strProcBodyLine = .Lines(.ProcBodyLine(.ProcOfLine(i, _
                                     vbext_pk_Proc), vbext_pk_Proc), 1)
                                    
                     End If
                 End If
             Next i
         End With
     Next vbComp
     
     wsList.Columns("A:B").Columns.AutoFit
 End Sub

