Sub Wijk_select_dashboards()
'
' Wijk_select_dashboards Macro
' Ik loop over m'n wijkselects heen, en print de dashboards.
'
'
Dim Jaar As String
Dim Kwartaal As String
Dim path As String
Dim Folder As String

Sheets("Chart_data").Select
Kwartaal = Range("AC4")
    Sheets("Wijkselectie").Select
    
    path = "Q:\Dashboards\" & "Newrapports" & "\" & "Wijkoverzichten"
    Folder = Dir(path, vbDirectory)
    If Folder = vbNullString Then
        MkDir path
    End If
    
    With ActiveSheet.PivotTables("Draaitabel3").PivotFields("WIJK_SELECT")
        .PivotItems("01_BINNEN").Visible = True
        .PivotItems("02_BUITEN").Visible = False
        .PivotItems("99_NIET").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    Sheets("Binnen-Buitendering").Select
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path & "\" & "Amsterdam binnen de ring - Kwartaalrapport " & Kwartaal & ".pdf", _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Sheets("Wijkselectie").Select
    With ActiveSheet.PivotTables("Draaitabel3").PivotFields("WIJK_SELECT")
        .PivotItems("02_BUITEN").Visible = True
        .PivotItems("01_BINNEN").Visible = False
        .PivotItems("99_NIET").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    Sheets("Binnen-Buitendering").Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path & "\" & "Amsterdam buiten de ring - Kwartaalrapport " & Kwartaal & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
    Sheets("Wijkselectie").Select
    With ActiveSheet.PivotTables("Draaitabel3").PivotFields("WIJK_SELECT")
        .PivotItems("01_BINNEN").Visible = True
        .PivotItems("02_BUITEN").Visible = True
        .PivotItems("99_NIET").Visible = True
        .PivotItems("(blank)").Visible = True
    End With
    Sheets("Geheel Amsterdam").Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path & "\" & "Geheel Amsterdam - Kwartaalrapport " & Kwartaal & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
    Sheets("Lijst wijken Jaar").Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path & "\" & "Lijst wijken op jaar - " & Kwartaal & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
    Sheets("Lijst wijken kwartaal").Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path & "\" & "Lijst wijken op kwartaal - " & Kwartaal & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
    Sheets("Subwijken tov vorig jaar").Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path & "\" & "Subwijken tov vorig jaar - " & Kwartaal & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
    Sheets("Subwijken tov vorig kwartaal").Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path & "\" & "Subwijken tov vorig kwartaal - " & Kwartaal & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
    Sheets("Subwijken tov vorig kwartaal").Select
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path & "\" & "Subwijken tov vorig kwartaal - " & Kwartaal & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
    Sheets("Wijkselectie").Select
    With ActiveSheet.PivotTables("Draaitabel3").PivotFields("PRIJSKLASSE")
        .PivotItems("TOT>1.000.000").Visible = True
        .PivotItems("TOT_#GEENTRPRS!").Visible = False
        .PivotItems("TOT__100.000").Visible = False
        .PivotItems("TOT__250.000").Visible = False
        .PivotItems("TOT__500.000").Visible = False
        .PivotItems("TOT__750.000").Visible = False
        .PivotItems("TOT_1.000.000").Visible = False
    End With
    
    path = "Q:\Dashboards\" & "Newrapports" & "\" & "Miljoenenrapportages"
    Folder = Dir(path, vbDirectory)
    If Folder = vbNullString Then
        MkDir path
    End If
    
    Sheets("Geheel-Milj").Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path & "\" & "Miljoenenrapportage Geheel Amsterdam - " & Kwartaal & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
    
    Sheets("Wijkselectie").Select
    With ActiveSheet.PivotTables("Draaitabel3").PivotFields("WIJK_SELECT")
        .PivotItems("01_BINNEN").Visible = True
        .PivotItems("02_BUITEN").Visible = False
        .PivotItems("99_NIET").Visible = False
        .PivotItems("(blank)").Visible = False
    End With

    Sheets("BinnenBuiten-Milj").Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path & "\" & "Miljoenenrapportage Binnen de ring - " & Kwartaal & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False

    Sheets("Wijkselectie").Select
    With ActiveSheet.PivotTables("Draaitabel3").PivotFields("WIJK_SELECT")
        .PivotItems("02_BUITEN").Visible = True
        .PivotItems("01_BINNEN").Visible = False
        .PivotItems("99_NIET").Visible = False
        .PivotItems("(blank)").Visible = False
    End With

    Sheets("BinnenBuiten-Milj").Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path & "\" & "Miljoenenrapportage Buiten de ring - " & Kwartaal & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
End Sub
