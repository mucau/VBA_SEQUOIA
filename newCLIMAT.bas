Attribute VB_Name = "Module18"
Sub newCLIMAT()
Dim Wb As Workbook: Set Wb = ActiveWorkbook

'Mois
Dim sh As Worksheet

For Each ws In Sheets
    ws.Activate
    
    ws.Cells(1, 2) = "Mois"
    ws.Cells(2, 2) = "Janv."
    ws.Cells(3, 2) = "Fev."
    ws.Cells(4, 2) = "Mars"
    ws.Cells(5, 2) = "Avril"
    ws.Cells(6, 2) = "Mai"
    ws.Cells(7, 2) = "Juin"
    ws.Cells(8, 2) = "Juil."
    ws.Cells(9, 2) = "Aout"
    ws.Cells(10, 2) = "Sep."
    ws.Cells(11, 2) = "Oct."
    ws.Cells(12, 2) = "Nov."
    ws.Cells(13, 2) = "Dec."
    
    If ws.Name = "METEOFRANCE_ombro" Or ws.Name = "AURELHY_ombro" Or ws.Name = "DRIAS_ombro" Then 'Perform the Excel action you wish (turn cell yellow below)
        ws.Cells(1, 8).Select
        
        ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
        With ActiveChart
            .HasTitle = False
            .SeriesCollection.NewSeries
            .FullSeriesCollection(1).Name = "=""Précipitation (mm)"""
            .FullSeriesCollection(1).Values = Range("F2:F13")
            .FullSeriesCollection(1).XValues = Range("B2:B13")
            .SeriesCollection.NewSeries
            .FullSeriesCollection(2).Name = "=""Température (°C)"""
            .FullSeriesCollection(2).Values = Range("D2:D13")
            .ChartType = xlColumnClustered
            .FullSeriesCollection(1).ChartType = xlColumnClustered
            .FullSeriesCollection(1).AxisGroup = 1
            .FullSeriesCollection(2).ChartType = xlLine
            .FullSeriesCollection(2).AxisGroup = 2
            .SetElement (msoElementPrimaryValueAxisShow)
            .SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Précipitation (mm)"
            .Axes(xlValue, xlSecondary).HasTitle = True
            .Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Température (°C)"
            .Axes(xlValue, xlSecondary).MaximumScale = (ActiveChart.Axes(xlValue).MaximumScale / 2)
        End With
    End If
    
    If ws.Name = "METEOFRANCE_etp" Or ws.Name = "AURELHY_etp" Or ws.Name = "DRIAS_etp" Then 'Perform the Excel action you wish (turn cell yellow below)
        ws.Cells(1, 11).Select
        
        ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
        With ActiveChart
            .HasTitle = False
            .SeriesCollection.NewSeries
            .FullSeriesCollection(1).Name = "=""P"""
            .FullSeriesCollection(1).Values = Range("C2:C13")
            .SeriesCollection.NewSeries
            .FullSeriesCollection(2).Name = "=""ETP"""
            .FullSeriesCollection(2).Values = Range("D2:D13")
            .SeriesCollection.NewSeries
            .FullSeriesCollection(3).Name = "=""P-ETP"""
            .FullSeriesCollection(3).Values = Range("F2:F13")
            .FullSeriesCollection(3).XValues = Range("B2:B13")
            .Axes(xlCategory).TickLabelPosition = xlLow
            .SetElement (msoElementLegendRight)
            .Legend.Position = xlBottom
        End With
    End If

Next ws


End Sub
