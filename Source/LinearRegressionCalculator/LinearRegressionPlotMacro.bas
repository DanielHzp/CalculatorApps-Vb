Attribute VB_Name = "Module2"
Option Explicit



'Macro that automates the creation of linear regression plots
'It is used internally by the userform backend
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
    Range("A2:B11").Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.SetSourceData Source:=Range("Sheet4!$A$2:$B$11")
    ActiveChart.ChartTitle.Select
    Selection.Delete
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    Selection.Caption = "y"
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.Caption = "x"
    ActiveChart.ChartArea.Select
    ActiveChart.PlotArea.Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).XValues = "=Sheet4!$A$2:$A$11"
    ActiveChart.FullSeriesCollection(2).Values = "=Sheet4!$C$2:$C$11"
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
    End With
    Selection.MarkerStyle = -4142
    ActiveChart.FullSeriesCollection(2).Smooth = True
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.SetElement (msoElementLegendRight)
    ActiveChart.PlotArea.Select
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).Name = "=""Experimental Data"""
    ActiveChart.FullSeriesCollection(2).Name = "=""Model prediction"""
End Sub
