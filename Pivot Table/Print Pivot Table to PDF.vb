Sub PrintDailyAmounts()
    With Main.PageSetup
        .CenterHeader = "Recovery Amounts"
        .Orientation = xlPortrait
        .PrintArea = PivotTables("tDailyAmounts").TableRange2.Address
        .Zoom = False
        .FitToPagesWide = 1
    End With
    
    Main.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=ActiveWorkbook.Path & "\" & "Recovery Amounts.pdf", _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
End Sub
