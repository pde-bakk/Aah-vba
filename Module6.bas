Sub All_macros()

Dim Start, Finish, TotalTime As Single
Dim hours, minutes, seconds As Long


If (MsgBox("Press Yes to run a timer on all the macros", vbYesNo)) = vbYes Then
    Start = Timer 'Set start time
    
    Call Subwijk_dashboards

    Call Wijk_dashboards

    Call Wijk_select_dashboards
    
    Finish = Timer 'Set end time
    TotalTime = Finish - Start
    
    hours = TotalTime / 3600
    TotalTime = TotalTime - (hours * 3600)
    minutes = TotalTime / 60
    TotalTime = TotalTime - (mins * 60)
    seconds = TotalTime
    
    MsgBox ("Macros ran for " & hours & ":" & minutes & ":" & seconds & ".")

Else
    Call Subwijk_dashboards

    Call Wijk_dashboards

    Call Wijk_select_dashboards
    End
End If


End Sub
