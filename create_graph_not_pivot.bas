Attribute VB_Name = "create_graph_not_pivot"
Option Explicit

'Inserting A Chart

Method 1:
Sub CreateChart()
'PURPOSE: Create a chart (chart dimensions are not required)

Dim rng As Range
Dim cht As Object

'Your data range for the chart
  Set rng = ActiveSheet.Range("A24:M27")

'Create a chart
  Set cht = ActiveSheet.Shapes.AddChart2

'Give chart some data
  cht.Chart.SetSourceData Source:=rng

'Determine the chart type
  cht.Chart.ChartType = xlXYScatterLines

End Sub

Method 2:
Sub CreateChart()
'PURPOSE: Create a chart (chart dimensions are required)

Dim rng As Range
Dim cht As ChartObject

'Your data range for the chart
  Set rng = ActiveSheet.Range("A24:M27")

'Create a chart
  Set cht = ActiveSheet.ChartObjects.Add( _
    Left:=ActiveCell.Left, _
    Width:=450, _
    Top:=ActiveCell.Top, _
    Height:=250)

'Give chart some data
  cht.Chart.SetSourceData Source:=rng

'Determine the chart type
  cht.Chart.ChartType = xlXYScatterLines
       
End Sub

'Looping Through Charts & Series
Sub LoopThroughCharts()
'PURPOSE: How to cycle through charts and chart series

Dim cht As ChartObject
Dim ser As Series

'Loop Through all charts on ActiveSheet
  For Each cht In ActiveSheet.ChartObjects

  Next cht

'Loop through all series in a chart
  For Each ser In grph.Chart.SeriesCollection

  Next ser
  
'Loop Through all series on Activesheet
  For Each cht In ActiveSheet.ChartObjects
    For Each ser In grph.Chart.SeriesCollection

    Next ser
  Next cht
  
End Sub

'Adding & Modifying A Chart Title
Sub AddChartTitle()
'PURPOSE: Add a title to a specific chart

Dim cht As ChartObject

Set cht = ActiveSheet.ChartObjects("Chart 1")

'Ensure chart has a title
  cht.Chart.HasTitle = True

'Change chart's title
  cht.Chart.ChartTitle.Text = "My Graph"

End Sub

Sub RepositionChartTitle()
'PURPOSE: Reposition a chart's title

Dim cht As ChartObject

Set cht = ActiveSheet.ChartObjects("Chart 1")

'Reposition title
  With cht.Chart.ChartTitle
    .Left = 100
    .Top = 50
  End With

End Sub

'Adding & Modifying A Graph Legend
Sub InsertChartLegend()

Dim cht As Chart

Set cht = ActiveSheet.ChartObjects("Chart 1").Chart

'Add Legend to the Right
  cht.SetElement (msoElementLegendRight)
  
'Add Legend to the Left
  cht.SetElement (msoElementLegendLeft)
  
'Add Legend to the Bottom
  cht.SetElement (msoElementLegendBottom)
  
'Add Legend to the Top
  cht.SetElement (msoElementLegendTop)
  
'Add Overlaying Legend to the Left
  cht.SetElement (msoElementLegendLeftOverlay)
  
'Add Overlaying Legend to the Right
  cht.SetElement (msoElementLegendRightOverlay)

End Sub

Sub DimensionChartLegend()

Dim lgd As Legend

Set lgd = ActiveSheet.ChartObjects("Chart 1").Chart.Legend

lgd.Left = 240.23
lgd.Top = 6.962
lgd.Width = 103.769
lgd.Height = 25.165

End Sub

'Adding Various Chart Attributes
Sub AddStuffToChart()

Dim cht As Chart

Set cht = ActiveSheet.ChartObjects("Chart 1").Chart

'Add X-axis
  cht.HasAxis(xlCategory, xlPrimary) = True '[Method #1]
  cht.SetElement (msoElementPrimaryCategoryAxisShow) '[Method #2]
  
'Add X-axis title
  cht.Axes(xlCategory, xlPrimary).HasTitle = True '[Method #1]
  cht.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis) '[Method #2]
  
'Add y-axis
  cht.HasAxis(xlValue, xlPrimary) = True '[Method #1]
  cht.SetElement (msoElementPrimaryValueAxisShow) '[Method #2]
  
'Add y-axis title
  cht.Axes(xlValue, xlPrimary).HasTitle = True '[Method #1]
  cht.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis) '[Method #2]
  
'Add Data Labels (Centered)
  cht.SetElement (msoElementDataLabelCenter)

'Add Major Gridlines
  cht.SetElement (msoElementPrimaryValueGridLinesMajor)
    
'Add Linear Trend Line
  cht.SeriesCollection(1).Trendlines.Add Type:=xlLinear
  
End Sub

'Modifying Various Chart Attributes
Sub ChangeChartFormatting()

Dim cht As Chart

Set cht = ActiveSheet.ChartObjects("Chart 1").Chart

'Adjust y-axis Scale
  cht.Axes(xlValue).MinimumScale = 40
  cht.Axes(xlValue).MaximumScale = 100

'Adjust x-axis Scale
  cht.Axes(xlCategory).MinimumScale = 1
  cht.Axes(xlCategory).MaximumScale = 10
  
'Adjust Bar Gap
  cht.ChartGroups(1).GapWidth = 60

'Format Font Size
  cht.ChartArea.Format.TextFrame2.TextRange.Font.Size = 12
  
'Format Font Type
  cht.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Arial"
  
'Make Font Bold
  cht.ChartArea.Format.TextFrame2.TextRange.Font.Bold = msoTrue
  
'Make Font Italicized
  cht.ChartArea.Format.TextFrame2.TextRange.Font.Italic = msoTrue

End Sub

'Removing Various Chart Attributes
Sub RemoveChartFormatting()

Dim cht As Chart

Set cht = ActiveSheet.ChartObjects("Chart 1").Chart

'Remove Chart Series
  cht.SeriesCollection(2).Delete

'Remove Gridlines
  cht.Axes(xlValue).MajorGridlines.Delete
  cht.Axes(xlValue).MinorGridlines.Delete
  
'Remove X-axis
  cht.Axes(xlCategory).Delete

'Remove Y-axis
  cht.Axes(xlValue).Delete

'Remove Legend
  cht.Legend.Delete

'Remove Title
  cht.ChartTitle.Delete
    
'Remove ChartArea border
  cht.ChartArea.Border.LineStyle = xlNone

'No background color fill
  cht.ChartArea.Format.Fill.Visible = msoFalse
  cht.PlotArea.Format.Fill.Visible = msoFalse

End Sub

'Change Your Colors
Sub ChangeChartColors()

Dim cht As Chart

Set cht = ActiveSheet.ChartObjects("Chart 1").Chart

'Change first bar chart series fill color
  cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(91, 155, 213)

'Change X-axis label color
  cht.Axes(xlCategory).TickLabels.Font.Color = RGB(91, 155, 213)

'Change Y-axis label color
  cht.Axes(xlValue).TickLabels.Font.Color = RGB(91, 155, 213)
  
'Change Plot Area border color
  cht.PlotArea.Format.Line.ForeColor.RGB = RGB(91, 155, 213)
  
'Change Major gridline color
  cht.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(91, 155, 213)

'Change Chart Title font color
  cht.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(91, 155, 213)

'No background color fill
  cht.ChartArea.Format.Fill.Visible = msoFalse
  cht.PlotArea.Format.Fill.Visible = msoFalse

End Sub


