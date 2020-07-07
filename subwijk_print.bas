Sub Subwijk_dashboards()
'
' Subwijk_dashboards Macro
'
Application.DisplayAlerts = False
Dim Kwartaal As String
Sheets("Chart_data").Select
Kwartaal = Range("AC4")

Sheets("Wijkselectie").Select
Dim l As Long
Dim cur As Long
Dim Subwijk As String
Dim naam As String
Dim path As String
Dim subpath As String
Dim Folder As String

l = ActiveSheet.PivotTables("Draaitabel3").PivotFields("SUBWIJK").PivotItems.Count

With ActiveSheet.PivotTables("Draaitabel3").PivotFields("SUBWIJK")
    .PivotItems(1).Visible = True
    For cur = 2 To l
        .PivotItems(cur).Visible = False
    Next cur
    
    For cur = 1 To l
        Subwijk = .PivotItems(cur).Name
        Debug.Print cur; Subwijk; Date
        
        Sheets("Subwijk").Select
        subpath = "Q:\Dashboards\" & "Newrapports" & "\Subwijken\"
        path = "Q:\Dashboards\" & "Newrapports"
        Folder = Dir(path, vbDirectory)
        If Folder = vbNullString Then
            MkDir path
        End If
        Folder = Dir(subpath, vbDirectory)
        If Folder = vbNullString Then
            MkDir subpath
        End If
        
        naam = subpath & Subwijk & " - Kwartaalrapport " & Kwartaal & ".pdf"
        Debug.Print "naam is " & naam
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=naam, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
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
