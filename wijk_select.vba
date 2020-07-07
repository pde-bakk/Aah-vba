Sub Wijk_select_dashboards()
'
' Wijk_select_dashboards Macro
' Ik loop over m'n wijkselects heen, en print de dashboards.
'

'
Dim Jaar As String
Dim Kwartaal As String
Jaar = "2020"
Kwartaal = "Q2"

    With ActiveSheet.PivotTables("Draaitabel3").PivotFields("WIJK_SELECT")
        .PivotItems("02_BUITEN").Visible = False
        .PivotItems("99_NIET").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    Sheets("Binnen-Buitendering").Select
    ActiveWindow.SmallScroll Down:=-12
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "Q:\Dashboards\Rapporten\Newrapports\Amsterdam binnen de ring - Kwartaalrapport " & Jaar & Kwartaal & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
    Sheets("Wijkselectie").Select
    With ActiveSheet.PivotTables("Draaitabel3").PivotFields("WIJK_SELECT")
        .PivotItems("01_BINNEN").Visible = False
        .PivotItems("02_BUITEN").Visible = True
    End With
    Sheets("Binnen-Buitendering").Select
    ActiveWindow.SmallScroll Down:=-12
    Range("M21").Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "Q:\Dashboards\Rapporten\Newrapports\Amsterdam buiten de ring - Kwartaalrapport " & Jaar & Kwartaal & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
    Sheets("Wijkselectie").Select
    With ActiveSheet.PivotTables("Draaitabel3").PivotFields("WIJK_SELECT")
        .PivotItems("01_BINNEN").Visible = True
        .PivotItems("99_NIET").Visible = True
        .PivotItems("(blank)").Visible = True
    End With
    Sheets("Geheel Amsterdam").Select
    ActiveWindow.SmallScroll Down:=-102
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "Q:\Dashboards\Rapporten\Newrapports\Geheel Amsterdam - Kwartaalrapport " & Jaar & Kwartaal & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
End Sub
