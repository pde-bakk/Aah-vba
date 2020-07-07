Sub Wijk_dashboards()
'
' Wijk_dashboards Macro
'
Application.DisplayAlerts = False
Dim Kwartaal As String
Sheets("Chart_data").Select
Kwartaal = Range("AC4")

Sheets("Wijkselectie").Select
Dim l As Long
Dim cur As Long
Dim Wijk As String
Dim naam As String

l = ActiveSheet.PivotTables("Draaitabel3").PivotFields("WIJK").PivotItems.Count

With ActiveSheet.PivotTables("Draaitabel3").PivotFields("WIJK")
    .PivotItems(1).Visible = True
    For cur = 2 To l
        .PivotItems(cur).Visible = False
    Next cur
    
    For cur = 1 To l
        Wijk = .PivotItems(cur).Name
        Debug.Print cur; Wijk
        
        Sheets("Wijk").Select
        naam = "Q:\Dashboards\Newrapports\Wijken\" & Wijk & " - Kwartaalrapport " & Kwartaal & ".pdf"
        Debug.Print "naam is " & naam
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:="Q:\Dashboards\Newrapports\Wijken\" & Wijk & " - Kwartaalrapport " & Kwartaal & ".pdf", _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
  '      Sheets("Wijk").Select
   '     ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:="Newrapports\Wijken\" & Wijk & " - Kwartaalrapport " & Kwartaal & ".pdf", _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("Wijkselectie").Select
        
        If cur + 1 <= l Then
            .PivotItems(cur + 1).Visible = True
            .PivotItems(cur).Visible = False
            Debug.Print .PivotItems(cur).Name & " is now False " & .PivotItems(cur + 1).Name & " is now True!"
        End If
        
        
    Next cur
    
    For cur = 1 To l
        Debug.Print cur; .PivotItems(cur).Name
        .PivotItems(cur).Visible = True
    Next cur
    
End With
Application.DisplayAlerts = True
End Sub
