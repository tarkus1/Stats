Attribute VB_Name = "StartStatsInAddIn"

Sub StartStats()
Attribute StartStats.VB_ProcData.VB_Invoke_Func = "S\n14"

    If MsgBox("start stats?", vbOKCancel) = vbOK Then
        ParticipantDate.Show
    Else
        Exit Sub
    End If
End Sub

