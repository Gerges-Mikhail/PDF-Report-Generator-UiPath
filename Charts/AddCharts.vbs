Sub AddCharts()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim chart1 As ChartObject
    Set chart1 = ws.ChartObjects.Add(Left:=200, Width:=230, Top:=5, Height:=225)
    chart1.Chart.ChartType = xlColumnClustered
    chart1.Chart.SetSourceData Source:=ws.Range("A1:B5")
    chart1.Chart.HasTitle = True
    chart1.Chart.ChartTitle.Text = "Sales Per Region"

    Dim chart2 As ChartObject
    Set chart2 = ws.ChartObjects.Add(Left:=10, Width:=400, Top:=250, Height:=225)
    chart2.Chart.ChartType = xlLine
    chart2.Chart.SetSourceData Source:=ws.Range("C1:D13")
    chart2.Chart.HasTitle = True
    chart2.Chart.ChartTitle.Text = "Monthly Sales"

End Sub