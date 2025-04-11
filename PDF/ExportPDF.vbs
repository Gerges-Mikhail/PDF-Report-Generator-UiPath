Sub ExportToPDF()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' Change to your sheet index or name
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:="M:\RPAUI\UiPathProjects\PDF Report Generator-UiPath\Output\SalesReport.pdf"
End Sub