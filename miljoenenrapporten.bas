Sub Miljoenen(ByVal Wijk As String, ByVal subpath As String)
'
' milli Macro
'
Dim naam As String

If Wijk = "Oud-Zuid" Or Wijk = "Centrum" Then
    ActiveSheet.PivotTables("Draaitabel3").PivotFields("PRIJSKLASSE").CurrentPage _
    = "(All)"
    With ActiveSheet.PivotTables("Draaitabel3").PivotFields("PRIJSKLASSE")
        .PivotItems("TOT_#GEENTRPRS!").Visible = False
        .PivotItems("TOT__100.000").Visible = False
        .PivotItems("TOT__250.000").Visible = False
        .PivotItems("TOT__500.000").Visible = False
        .PivotItems("TOT__750.000").Visible = False
        .PivotItems("TOT_1.000.000").Visible = False
    End With

    naam = subpath & "\" & Wijk & " Miljoenenrapportage - Kwartaalrapport " & Kwartaal & ".pdf"
    Debug.Print "miljoenen naam: " & naam
    Sheets("Wijk-Miljoenen").Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=naam, _
    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
    :=False, OpenAfterPublish:=False
        
    Sheets("Wijkselectie").Select
    ActiveSheet.PivotTables("Draaitabel3").PivotFields("PRIJSKLASSE").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Draaitabel3").PivotFields("PRIJSKLASSE")
        .PivotItems("TOT_#GEENTRPRS!").Visible = True
        .PivotItems("TOT__100.000").Visible = True
        .PivotItems("TOT__250.000").Visible = True
        .PivotItems("TOT__500.000").Visible = True
        .PivotItems("TOT__750.000").Visible = True
        .PivotItems("TOT_1.000.000").Visible = True
    End With
     
End If
End Sub
