Attribute VB_Name = "modReports"
Sub btn_PDF_Transfer()
    Dim nameFiles As String
    Dim ws As Worksheet
    Set ws = Sheets("DASHBOARD")
    
    
    nameFiles = ThisWorkbook.Path & "\Financial_Report_" & Format(Date, "dd_mm_yyyy") & ".pdf"
    
    
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.2)
        .RightMargin = Application.InchesToPoints(0.2)
        .TopMargin = Application.InchesToPoints(0.2)
        .BottomMargin = Application.InchesToPoints(0.2)
    End With
    
    
    On Error Resume Next
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=nameFiles, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    
    If Err.Number = 0 Then
        MsgBox "The report has been successfully generated in landscape mode: " & vbCrLf & nameFiles, vbInformation
    Else
        MsgBox "Error: Could not generate PDF. Please check if the file is open.", vbCritical
    End If
    On Error GoTo 0
End Sub


