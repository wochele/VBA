Sub PrintTableToPdf()
    With TableToPdf.PageSetup
        .CenterHeader = "Table to PDF"
        .Orientation = xlPortrait ' or xlLandscape
        .PrintArea = Range("tUsers").Address
        .PrintTitleRows = TableToPdf.Rows(3).Address 'Use row the headers are on to print the headers
        .Zoom = False  'This is set to false so the .FitToPagesWide and .FitToPagesTall properties are used
        .FitToPagesWide = 1  'Force to print one page wide
    End With
    
    TableToPdf.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=ActiveWorkbook.Path & "\" & "TableToPdf.pdf", _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
End Sub
