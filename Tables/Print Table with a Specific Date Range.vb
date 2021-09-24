Sub PrintDateRange()
    'Filter payments table between dates
    Data.ListObjects("tPayments").Range.AutoFilter Field:=1, Criteria1:=">=" & Range("StartDate"), Operator:=xlAnd, Criteria2:="<=" & Range("EndDate")
    
    With Data.PageSetup
        .CenterHeader = "Recoveries for " & Main.Range("P3").value & " : Total(" & FormatCurrency(Data.Range("V2").value) & ")"
        .Orientation = xlPortrait
        .PrintArea = Data.Range("tPayments").Address
        .PrintTitleRows = Data.Rows(3).Address
        .Zoom = False
        .FitToPagesWide = 1
    End With
    
    Data.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=ActiveWorkbook.Path & "\" & "Recovery.pdf", _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    'Turn off Autofilter
    Data.ListObjects("tPayments").Range.AutoFilter
End Sub
