Function Col_lett(ByVal ColumnNumber As Integer)
Col_lett = Replace(Replace(Cells(1, ColumnNumber).Address, "1", ""), "$", "")
End Function
Sub Weather()
Dim Width As Single, Height As Single, NumWide As Long

'''''''''''''''''''''''''''Trivial assuming no changes to the worksheet: '''''''''''''''''''''''''''

'    Columns("Q:Q").Font.Bold = True
'    Columns("Q:Q").EntireColumn.AutoFit
'    Range("Q4").FormulaR1C1 = "Average Temp"
'    Range("Q5").FormulaR1C1 = "=AVERAGE(RC[-12]:RC[-1])"
'    Range("Q5").AutoFill Destination:=Range("Q5:Q120"), Type:=xlFillDefault
'
'    Rows("121:121").Font.Bold = True
'    Range("C121").FormulaR1C1 = "AVERAGE TEMPERATURE"
'    Range("E121").FormulaR1C1 = "=AVERAGE(R[-116]C:R[-1]C)"
'    Range("E121").AutoFill Destination:=Range("E121:P121"), Type:=xlFillDefault
'
'    Range("Q126").FormulaR1C1 = "Total PRECIPITATION"
'    Range("Q127").FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
'    Range("Q127").AutoFill Destination:=Range("Q127:Q266"), Type:=xlFillDefault
'
'    Rows("267:267").Font.Bold = True
'    Range("C267").FormulaR1C1 = "Average"
'    Range("E267").FormulaR1C1 = "=AVERAGE(R[-140]C:R[-1]C)"
'    Range("E267").AutoFill Destination:=Range("E267:P267"), Type:=xlFillDefault
    
'''''''''''''''''''''''''''''''''''''Non-trivial Algorithm''''''''''''''''''''''''''''''''''''''''''
    
'However, the following case assumptions are still made:
'1) "STA_ID" exists but not neccessarily as the aforementioned string
'2) Both temperature and precipation data are seperated by at least one empty row
'3) No further changes have been made including calculations and graph creations


''''''''''''''TEMPERATURE'''''''''''''''

    Range("A1").Select
    city_name = Cells.Find(What:="*STA_NAME", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Offset(1, 0).Value
    country_name = Cells.Find(What:="*COUNTRY", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Offset(1, 0).Value
        
        
    Cells.Find(What:="*STA_ID", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    STA_ID_row = ActiveCell.Row
    temp_last_row = ActiveCell.End(xlDown).Row
    
    type_column = Cells.Find(What:="*TYPE", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Column
        
    year_column = Cells.Find(What:="*Year", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Column
        
    january_column = Cells.Find(What:="*JAN", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Column
    
    december_column = Cells.Find(What:="*DEC", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Column

''''Columns''''

    For i = 1 To 4
    Columns(Col_lett(december_column + 1) & ":" & Col_lett(december_column + 1)).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next i
    
''''''''ANNUAL'''''''''''
    
    Columns(Col_lett(december_column + 1) & ":" & Col_lett(december_column + 1)).Font.Bold = True
    Range(Col_lett(december_column + 1) & STA_ID_row).FormulaR1C1 = "AVERAGE ANNUAL TEMP"
    Range(Col_lett(december_column + 1) & (STA_ID_row + 1)).FormulaR1C1 = "=AVERAGE(RC[" & -(december_column + 1 - january_column) & "]:RC[-1])"
    Range(Col_lett(december_column + 1) & (STA_ID_row + 1)).AutoFill Destination:=Range(Col_lett(december_column + 1) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 1) & temp_last_row), Type:=xlFillDefault
    
''''''''WINTER'''''''''''


    Columns(Col_lett(december_column + 2) & ":" & Col_lett(december_column + 2)).Font.Bold = True
    Range(Col_lett(december_column + 2) & STA_ID_row).FormulaR1C1 = "AVERAGE WINTER TEMP"
    Range(Col_lett(december_column + 2) & (STA_ID_row + 1)).FormulaR1C1 = "=AVERAGE(RC[-2],RC[" & -(december_column + 2 - january_column) & "]:RC[" & -(december_column + 1 - january_column) & "])"
    Range(Col_lett(december_column + 2) & (STA_ID_row + 1)).AutoFill Destination:=Range(Col_lett(december_column + 2) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 2) & temp_last_row), Type:=xlFillDefault


''''''''SPRING'''''''''''
    
    Columns(Col_lett(december_column + 3) & ":" & Col_lett(december_column + 3)).Font.Bold = True
    Range(Col_lett(december_column + 3) & STA_ID_row).FormulaR1C1 = "AVERAGE SPRING TEMP"
    Range(Col_lett(december_column + 3) & STA_ID_row + 1).FormulaR1C1 = "=AVERAGE(RC[" & -(december_column + 1 - january_column) & "]:RC[" & -(december_column - (january_column + 1)) & "])"
    Range(Col_lett(december_column + 3) & (STA_ID_row + 1)).AutoFill Destination:=Range(Col_lett(december_column + 3) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 3) & temp_last_row), Type:=xlFillDefault


''''''''SUMMER'''''''''''


    Columns(Col_lett(december_column + 4) & ":" & Col_lett(december_column + 4)).Font.Bold = True
    Range(Col_lett(december_column + 4) & STA_ID_row).FormulaR1C1 = "AVERAGE SUMMER TEMP"
    Range(Col_lett(december_column + 4) & STA_ID_row + 1).FormulaR1C1 = "=AVERAGE(RC[" & -(december_column - (january_column + 1)) & "]:RC[" & -(december_column - (january_column + 3)) & "])"
    Range(Col_lett(december_column + 4) & (STA_ID_row + 1)).AutoFill Destination:=Range(Col_lett(december_column + 4) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 4) & temp_last_row), Type:=xlFillDefault


'''''''''FALL''''''''''''
    
    Columns(Col_lett(december_column + 5) & ":" & Col_lett(december_column + 5)).Font.Bold = True
    Range(Col_lett(december_column + 5) & STA_ID_row).FormulaR1C1 = "AVERAGE FALL TEMP"
    Range(Col_lett(december_column + 5) & STA_ID_row + 1).FormulaR1C1 = "=AVERAGE(RC[" & -(december_column - (january_column + 3)) & "]:RC[" & -(december_column - (january_column + 5)) & "])"
    Range(Col_lett(december_column + 5) & (STA_ID_row + 1)).AutoFill Destination:=Range(Col_lett(december_column + 5) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 5) & temp_last_row), Type:=xlFillDefault
    
    
    
    Rows(temp_last_row + 1 & ":" & temp_last_row + 1).Font.Bold = True
    Range(Col_lett(type_column) & (temp_last_row + 1)).FormulaR1C1 = "AVERAGE"
    Range(Col_lett(january_column) & (temp_last_row + 1)).FormulaR1C1 = "=AVERAGE(R[" & -(temp_last_row - STA_ID_row) & "]C:R[-1]C)"
    Range(Col_lett(january_column) & (temp_last_row + 1)).AutoFill Destination:=Range(Col_lett(january_column) & (temp_last_row + 1) & ":" & Col_lett(december_column) & temp_last_row + 1), Type:=xlFillDefault
    
    Temp_first_year = Range(Col_lett(year_column) & (STA_ID_row + 1)).Value
    Temp_last_year = Range(Col_lett(year_column) & (temp_last_row)).Value

'''''''''''''Chart 1'''''''''''''''''''''

    Range(Col_lett(december_column + 1) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 1) & (temp_last_row) & "," & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & (temp_last_row)).Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(year_column) & "$" & (temp_last_row)
    ActiveChart.FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 1) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(december_column + 1) & "$" & (temp_last_row)
    Range("Sheet3!B6").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 1) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 1) & (temp_last_row) & ",Sheet1!" & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & (temp_last_row) & ")"
    
    With ActiveChart
        .ChartArea.Font.Name = "Arial"
        .ChartArea.Font.Color = RGB(31, 78, 121)
        .ChartTitle.Text = city_name & " " & country_name & " Average Annual Temp " & Range(Col_lett(year_column) & (STA_ID_row + 1)).Value & "-" & Range(Col_lett(year_column) & (temp_last_row)).Value & " Figure 1"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (years)"
        .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        
        .SetElement (msoElementPrimaryValueGridLinesNone)
        .SetElement (msoElementPrimaryCategoryGridLinesNone)
        
        .FullSeriesCollection(1).MarkerForegroundColorIndex = -4142
        .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
        .FullSeriesCollection(1).MarkerSize = 4
        .FullSeriesCollection(1).MarkerStyle = 8
        .FullSeriesCollection(1).Trendlines.Add
        .FullSeriesCollection(1).Trendlines(1).DisplayEquation = True
        .FullSeriesCollection(1).Trendlines(1).DisplayRSquared = True
        
        .Axes(xlCategory).Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
        .Axes(xlValue).Select
         With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
    End With
    
'''''''''''''Chart 2''''''''''''''''''''''
    
    Range(Col_lett(year_column) & (STA_ID_row)).Select
    y_1920 = Cells.Find(What:="1920", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row
        
    Range(Col_lett(december_column + 1) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 1) & (y_1920) & "," & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & (y_1920)).Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(year_column) & "$" & (y_1920)
    ActiveChart.FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 1) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(december_column + 1) & "$" & (y_1920)
    Range("Sheet3!E6").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 1) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 1) & (y_1920) & ",Sheet1!" & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & (y_1920) & ")"


    With ActiveChart
        .ChartArea.Font.Name = "Arial"
        .ChartArea.Font.Color = RGB(31, 78, 121)
        .ChartTitle.Text = city_name & " " & country_name & " Average Annual Temp " & Range(Col_lett(year_column) & (STA_ID_row + 1)).Value & "-1920 Figure 2"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (years)"
        .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        
        .SetElement (msoElementPrimaryValueGridLinesNone)
        .SetElement (msoElementPrimaryCategoryGridLinesNone)
        
        .FullSeriesCollection(1).MarkerForegroundColorIndex = -4142
        .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
        .FullSeriesCollection(1).MarkerSize = 4
        .FullSeriesCollection(1).MarkerStyle = 8
        .FullSeriesCollection(1).Trendlines.Add
        .FullSeriesCollection(1).Trendlines(1).DisplayEquation = True
        .FullSeriesCollection(1).Trendlines(1).DisplayRSquared = True
        
        .Axes(xlCategory).Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
        .Axes(xlValue).Select
         With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
    End With
    
'''''''''''''Chart 3''''''''''''''''''''''
    
    Range(Col_lett(year_column) & (STA_ID_row)).Select
    y_1950 = Cells.Find(What:="1950", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row
        
    Range(Col_lett(december_column + 1) & (y_1920) & ":" & Col_lett(december_column + 1) & (y_1950) & "," & Col_lett(year_column) & (y_1920) & ":" & Col_lett(year_column) & (y_1950)).Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (y_1920) & ":" & "$" & Col_lett(year_column) & "$" & (y_1950)
    ActiveChart.FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 1) & "$" & (y_1920) & ":" & "$" & Col_lett(december_column + 1) & "$" & (y_1950)
    Range("Sheet3!H6").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 1) & (y_1920) & ":" & Col_lett(december_column + 1) & (y_1950) & ",Sheet1!" & Col_lett(year_column) & (y_1920) & ":" & Col_lett(year_column) & (y_1950) & ")"


    With ActiveChart
        .ChartArea.Font.Name = "Arial"
        .ChartArea.Font.Color = RGB(31, 78, 121)
        .ChartTitle.Text = city_name & " " & country_name & " Average Annual Temp 1920-1950 Figure 3"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (years)"
        .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        
        .SetElement (msoElementPrimaryValueGridLinesNone)
        .SetElement (msoElementPrimaryCategoryGridLinesNone)
        
        .FullSeriesCollection(1).MarkerForegroundColorIndex = -4142
        .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
        .FullSeriesCollection(1).MarkerSize = 4
        .FullSeriesCollection(1).MarkerStyle = 8
        .FullSeriesCollection(1).Trendlines.Add
        .FullSeriesCollection(1).Trendlines(1).DisplayEquation = True
        .FullSeriesCollection(1).Trendlines(1).DisplayRSquared = True
        
        .Axes(xlCategory).Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
        .Axes(xlValue).Select
         With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
    End With


''''''''''''''Chart 4''''''''''''''''''''''

    Range(Col_lett(december_column + 1) & (y_1950) & ":" & Col_lett(december_column + 1) & (temp_last_row) & "," & Col_lett(year_column) & (y_1950) & ":" & Col_lett(year_column) & (temp_last_row)).Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (y_1950) & ":" & "$" & Col_lett(year_column) & "$" & (temp_last_row)
    ActiveChart.FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 1) & "$" & (y_1950) & ":" & "$" & Col_lett(december_column + 1) & "$" & (temp_last_row)
    Range("Sheet3!K6").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 1) & (y_1950) & ":" & Col_lett(december_column + 1) & (temp_last_row) & ",Sheet1!" & Col_lett(year_column) & (y_1950) & ":" & Col_lett(year_column) & (temp_last_row) & ")"


    With ActiveChart
        .ChartArea.Font.Name = "Arial"
        .ChartArea.Font.Color = RGB(31, 78, 121)
        .ChartTitle.Text = city_name & " " & country_name & " Average Annual Temp 1950-" & Range(Col_lett(year_column) & (temp_last_row)).Value & " Figure 4"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (years)"
        .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        
        .SetElement (msoElementPrimaryValueGridLinesNone)
        .SetElement (msoElementPrimaryCategoryGridLinesNone)
        
        .FullSeriesCollection(1).MarkerForegroundColorIndex = -4142
        .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
        .FullSeriesCollection(1).MarkerSize = 4
        .FullSeriesCollection(1).MarkerStyle = 8
        .FullSeriesCollection(1).Trendlines.Add
        .FullSeriesCollection(1).Trendlines(1).DisplayEquation = True
        .FullSeriesCollection(1).Trendlines(1).DisplayRSquared = True
        
        .Axes(xlCategory).Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
        .Axes(xlValue).Select
         With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
    End With
    

''''''''''''''Chart 5''''''''''''''''''''''

    Range(Col_lett(december_column + 2) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 2) & (temp_last_row) & "," & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & (temp_last_row)).Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(year_column) & "$" & (temp_last_row)
    ActiveChart.FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 2) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(december_column + 2) & "$" & (temp_last_row)
    Range("Sheet3!B7").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 2) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 2) & (temp_last_row) & ",Sheet1!" & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & (temp_last_row) & ")"
    
    With ActiveChart
        .ChartArea.Font.Name = "Arial"
        .ChartArea.Font.Color = RGB(31, 78, 121)
        .ChartTitle.Text = city_name & " " & country_name & " Average Winter(DJF) Temp " & Range(Col_lett(year_column) & (STA_ID_row + 1)).Value & "-" & Range(Col_lett(year_column) & (temp_last_row)).Value & " Figure 5"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (years)"
        .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        
        .SetElement (msoElementPrimaryValueGridLinesNone)
        .SetElement (msoElementPrimaryCategoryGridLinesNone)
        
        .FullSeriesCollection(1).MarkerForegroundColorIndex = -4142
        .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
        .FullSeriesCollection(1).MarkerSize = 4
        .FullSeriesCollection(1).MarkerStyle = 8
        .FullSeriesCollection(1).Trendlines.Add
        .FullSeriesCollection(1).Trendlines(1).DisplayEquation = True
        .FullSeriesCollection(1).Trendlines(1).DisplayRSquared = True
        
        .Axes(xlCategory).Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
        .Axes(xlValue).Select
         With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
    End With


''''''''''''''Chart 6''''''''''''''''''''''

    Range(Col_lett(december_column + 3) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 3) & (temp_last_row) & "," & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & (temp_last_row)).Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(year_column) & "$" & (temp_last_row)
    ActiveChart.FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 3) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(december_column + 3) & "$" & (temp_last_row)
    Range("Sheet3!B8").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 3) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 3) & (temp_last_row) & ",Sheet1!" & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & (temp_last_row) & ")"
    
    With ActiveChart
        .ChartArea.Font.Name = "Arial"
        .ChartArea.Font.Color = RGB(31, 78, 121)
        .ChartTitle.Text = city_name & " " & country_name & " Average Spring(MAM) Temp " & Range(Col_lett(year_column) & (STA_ID_row + 1)).Value & "-" & Range(Col_lett(year_column) & (temp_last_row)).Value & " Figure 6"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (years)"
        .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        
        .SetElement (msoElementPrimaryValueGridLinesNone)
        .SetElement (msoElementPrimaryCategoryGridLinesNone)
        
        .FullSeriesCollection(1).MarkerForegroundColorIndex = -4142
        .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
        .FullSeriesCollection(1).MarkerSize = 4
        .FullSeriesCollection(1).MarkerStyle = 8
        .FullSeriesCollection(1).Trendlines.Add
        .FullSeriesCollection(1).Trendlines(1).DisplayEquation = True
        .FullSeriesCollection(1).Trendlines(1).DisplayRSquared = True
        
        .Axes(xlCategory).Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
        .Axes(xlValue).Select
         With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
    End With


''''''''''''''Chart 7''''''''''''''''''''''

    Range(Col_lett(december_column + 4) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 4) & (temp_last_row) & "," & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & (temp_last_row)).Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(year_column) & "$" & (temp_last_row)
    ActiveChart.FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 4) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(december_column + 4) & "$" & (temp_last_row)
    Range("Sheet3!B9").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 4) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 4) & (temp_last_row) & ",Sheet1!" & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & (temp_last_row) & ")"
    
    With ActiveChart
        .ChartArea.Font.Name = "Arial"
        .ChartArea.Font.Color = RGB(31, 78, 121)
        .ChartTitle.Text = city_name & " " & country_name & " Average Summer(JJA) Temp " & Range(Col_lett(year_column) & (STA_ID_row + 1)).Value & "-" & Range(Col_lett(year_column) & (temp_last_row)).Value & " Figure 7"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (years)"
        .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        
        .SetElement (msoElementPrimaryValueGridLinesNone)
        .SetElement (msoElementPrimaryCategoryGridLinesNone)
        
        .FullSeriesCollection(1).MarkerForegroundColorIndex = -4142
        .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
        .FullSeriesCollection(1).MarkerSize = 4
        .FullSeriesCollection(1).MarkerStyle = 8
        .FullSeriesCollection(1).Trendlines.Add
        .FullSeriesCollection(1).Trendlines(1).DisplayEquation = True
        .FullSeriesCollection(1).Trendlines(1).DisplayRSquared = True
        
        .Axes(xlCategory).Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
        .Axes(xlValue).Select
         With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
    End With


''''''''''''''Chart 8''''''''''''''''''''''

    Range(Col_lett(december_column + 5) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 5) & (temp_last_row) & "," & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & (temp_last_row)).Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(year_column) & "$" & (temp_last_row)
    ActiveChart.FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 5) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(december_column + 5) & "$" & (temp_last_row)
    Range("Sheet3!B10").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 5) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 5) & (temp_last_row) & ",Sheet1!" & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & (temp_last_row) & ")"
    
    With ActiveChart
        .ChartArea.Font.Name = "Arial"
        .ChartArea.Font.Color = RGB(31, 78, 121)
        .ChartTitle.Text = city_name & " " & country_name & " Average Fall(SON) Temp " & Range(Col_lett(year_column) & (STA_ID_row + 1)).Value & "-" & Range(Col_lett(year_column) & (temp_last_row)).Value & " Figure 8"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (years)"
        .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        
        .SetElement (msoElementPrimaryValueGridLinesNone)
        .SetElement (msoElementPrimaryCategoryGridLinesNone)
        
        .FullSeriesCollection(1).MarkerForegroundColorIndex = -4142
        .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
        .FullSeriesCollection(1).MarkerSize = 4
        .FullSeriesCollection(1).MarkerStyle = 8
        .FullSeriesCollection(1).Trendlines.Add
        .FullSeriesCollection(1).Trendlines(1).DisplayEquation = True
        .FullSeriesCollection(1).Trendlines(1).DisplayRSquared = True
        
        .Axes(xlCategory).Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
        .Axes(xlValue).Select
         With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
    End With


''''''''''''''PRECIPITATION'''''''''''''''''


    Range("A" & temp_last_row).Select
    Cells.Find(What:="*STA_ID", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    STA_ID_row = ActiveCell.Row
    prcp_last_row = ActiveCell.End(xlDown).Row


    type_column = Cells.Find(What:="*TYPE", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Column


    january_column = Cells.Find(What:="*JAN", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Column


    december_column = Cells.Find(What:="*DEC", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Column


    Columns(Col_lett(december_column + 1) & ":" & Col_lett(december_column + 1)).Font.Bold = True
    Range(Col_lett(december_column + 1) & STA_ID_row).FormulaR1C1 = "TOTAL PRECIPITATION"
    Range(Col_lett(december_column + 1) & (STA_ID_row + 1)).FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    Range(Col_lett(december_column + 1) & (STA_ID_row + 1)).AutoFill Destination:=Range(Col_lett(december_column + 1) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 1) & prcp_last_row), Type:=xlFillDefault


 ''''''''WINTER'''''''''''


    Range(Col_lett(december_column + 2) & STA_ID_row).FormulaR1C1 = "TOTAL WINTER PRECIPITATION"
    Range(Col_lett(december_column + 2) & (STA_ID_row + 1)).FormulaR1C1 = "=SUM(RC[-2],RC[" & -(december_column + 2 - january_column) & "]:RC[" & -(december_column + 1 - january_column) & "])"
    Range(Col_lett(december_column + 2) & (STA_ID_row + 1)).AutoFill Destination:=Range(Col_lett(december_column + 2) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 2) & prcp_last_row), Type:=xlFillDefault


''''''''SPRING'''''''''''
    
    Range(Col_lett(december_column + 3) & STA_ID_row).FormulaR1C1 = "TOTAL SPRING PRECIPITATION"
    Range(Col_lett(december_column + 3) & STA_ID_row + 1).FormulaR1C1 = "=SUM(RC[" & -(december_column + 1 - january_column) & "]:RC[" & -(december_column - (january_column + 1)) & "])"
    Range(Col_lett(december_column + 3) & (STA_ID_row + 1)).AutoFill Destination:=Range(Col_lett(december_column + 3) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 3) & prcp_last_row), Type:=xlFillDefault


''''''''SUMMER'''''''''''


    Range(Col_lett(december_column + 4) & STA_ID_row).FormulaR1C1 = "TOTAL SUMMER PRECIPITATION"
    Range(Col_lett(december_column + 4) & STA_ID_row + 1).FormulaR1C1 = "=SUM(RC[" & -(december_column - (january_column + 1)) & "]:RC[" & -(december_column - (january_column + 3)) & "])"
    Range(Col_lett(december_column + 4) & (STA_ID_row + 1)).AutoFill Destination:=Range(Col_lett(december_column + 4) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 4) & prcp_last_row), Type:=xlFillDefault


'''''''''FALL''''''''''''
    
    Range(Col_lett(december_column + 5) & STA_ID_row).FormulaR1C1 = "TOTAL FALL PRECIPITATION"
    Range(Col_lett(december_column + 5) & STA_ID_row + 1).FormulaR1C1 = "=SUM(RC[" & -(december_column - (january_column + 3)) & "]:RC[" & -(december_column - (january_column + 5)) & "])"
    Range(Col_lett(december_column + 5) & (STA_ID_row + 1)).AutoFill Destination:=Range(Col_lett(december_column + 5) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 5) & prcp_last_row), Type:=xlFillDefault
       
    
    Rows(prcp_last_row + 1 & ":" & prcp_last_row + 1).Font.Bold = True
    Range(Col_lett(type_column) & (prcp_last_row + 1)).FormulaR1C1 = "AVERAGE"
    Range(Col_lett(january_column) & (prcp_last_row + 1)).FormulaR1C1 = "=AVERAGE(R[" & -(prcp_last_row - STA_ID_row) & "]C:R[-1]C)"
    Range(Col_lett(january_column) & (prcp_last_row + 1)).AutoFill Destination:=Range(Col_lett(january_column) & (prcp_last_row + 1) & ":" & Col_lett(december_column + 1) & prcp_last_row + 1), Type:=xlFillDefault
    Range("Sheet3!E18").Value = Range(Col_lett(december_column + 1) & prcp_last_row + 1).Value

    Prcp_first_year = Range(Col_lett(year_column) & (STA_ID_row + 1)).Value
    Prcp_last_year = Range(Col_lett(year_column) & (prcp_last_row)).Value

'''''''Chart 9''''''''''''''

    Range(Col_lett(december_column + 1) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 1) & (prcp_last_row) & "," & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & (prcp_last_row)).Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(year_column) & "$" & (prcp_last_row)
    ActiveChart.FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 1) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(december_column + 1) & "$" & (prcp_last_row)
    Range("Sheet3!B18").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 1) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 1) & (prcp_last_row) & ",Sheet1!" & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & (prcp_last_row) & ")"


    With ActiveChart
        .ChartArea.Font.Name = "Arial"
        .ChartArea.Font.Color = RGB(31, 78, 121)
        .ChartTitle.Text = city_name & " " & country_name & " Total Annual Precipitation " & Range(Col_lett(year_column) & (STA_ID_row + 1)).Value & "-" & Range(Col_lett(year_column) & (prcp_last_row)).Value & " Figure 9"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (years)"
        .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside


        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Total Precipitation (mm)"
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 10
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside


        .SetElement (msoElementPrimaryValueGridLinesNone)
        .SetElement (msoElementPrimaryCategoryGridLinesNone)


        .FullSeriesCollection(1).MarkerForegroundColorIndex = -4142
        .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
        .FullSeriesCollection(1).MarkerSize = 4
        .FullSeriesCollection(1).MarkerStyle = 8
        .FullSeriesCollection(1).Trendlines.Add
        .FullSeriesCollection(1).Trendlines(1).DisplayEquation = True
        .FullSeriesCollection(1).Trendlines(1).DisplayRSquared = True


        .Axes(xlCategory).Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
        .Axes(xlValue).Select
         With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.5
            .Transparency = 0
        End With
    End With


''''''''''''''''Auto-Fit Column''''''''''''''''''
    
    Columns(Col_lett(december_column + 1) & ":" & Col_lett(december_column + 1)).EntireColumn.AutoFit
    Columns(Col_lett(december_column + 2) & ":" & Col_lett(december_column + 2)).EntireColumn.AutoFit
    Columns(Col_lett(december_column + 3) & ":" & Col_lett(december_column + 3)).EntireColumn.AutoFit
    Columns(Col_lett(december_column + 4) & ":" & Col_lett(december_column + 4)).EntireColumn.AutoFit
    Columns(Col_lett(december_column + 5) & ":" & Col_lett(december_column + 5)).EntireColumn.AutoFit
    
    
    Chart_count = ActiveSheet.ChartObjects.Count
    Width = 480
    Height = 280
    NumWide = (Chart_count / 3)

    For Chart_index = 1 To Chart_count
        With ActiveSheet.ChartObjects(Chart_index)
            .Width = Width
            .Height = Height
            .Left = ((Chart_index - 1) Mod NumWide) * Width
            .Top = Int((Chart_index - 1) / NumWide) * Height
        End With
    Next
    For Chart_index = 1 To Chart_count / 3
        ActiveSheet.ChartObjects(Chart_index).Top = 620
    Next
    For Chart_index = Chart_count / 3 + 1 To ((2 * Chart_count) / 3)
        ActiveSheet.ChartObjects(Chart_index).Top = 900
    Next
    For Chart_index = ((2 * Chart_count) / 3) + 1 To Chart_count
        ActiveSheet.ChartObjects(Chart_index).Top = 1180
    Next
    
    Sheets(ActiveSheet.Name).Name = "Main_Data"
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes.Range(Array("Chart 1", "Chart 2", "Chart 3", "Chart 4", _
        "Chart 5", "Chart 6", "Chart 7", "Chart 8", "Chart 9")).Select
    Selection.Cut
    Sheets("Sheet2").Select
    Range("A3").Select
    ActiveSheet.Paste
    Sheets(ActiveSheet.Name).Name = "Graphs"
    
    Sheets("Sheet3").Select
    Sheets(ActiveSheet.Name).Name = "Summary_Tables"
    
    Range("A2:L2").HorizontalAlignment = xlCenter
    Range("A2:L2").Merge
    Range("A2:L2").FormulaR1C1 = "Average Temperature Change"
    
    Range("A6").FormulaR1C1 = "All Seasons"
    Range("A7").FormulaR1C1 = "Winter"
    Range("A8").FormulaR1C1 = "Spring"
    Range("A9").FormulaR1C1 = "Summer"
    Range("A10").FormulaR1C1 = "Fall"
    
    Range("B4:C4").HorizontalAlignment = xlCenter
    Range("B4:C4").Merge
    Range("B4:C4").FormulaR1C1 = Temp_first_year & "-" & Temp_last_year
    Range("B5").FormulaR1C1 = Chr(176) & "C/year"
    Range("C5").FormulaR1C1 = "Total " & ChrW(8710) & "T (" & Chr(176) & "C)"
    Range("C6").FormulaR1C1 = "=RC[-1]*(" & Temp_last_year & "-" & Temp_first_year & ")"
    Range("C6").AutoFill Destination:=Range("C6:C10"), Type:=xlFillDefault
        
    Range("E4:F4").HorizontalAlignment = xlCenter
    Range("E4:F4").Merge
    Range("E4:F4").FormulaR1C1 = Temp_first_year & "-1920"
    Range("E5").FormulaR1C1 = Chr(176) & "C/year"
    Range("F5").FormulaR1C1 = "Total " & ChrW(8710) & "T (" & Chr(176) & "C)"
    Range("F6").FormulaR1C1 = "=RC[-1]*(1920-" & Temp_first_year & ")"
    
    Range("H4:I4").HorizontalAlignment = xlCenter
    Range("H4:I4").Merge
    Range("H4:I4").FormulaR1C1 = "1920-1950"
    Range("H5").FormulaR1C1 = Chr(176) & "C/year"
    Range("I5").FormulaR1C1 = "Total " & ChrW(8710) & "T (" & Chr(176) & "C)"
    Range("I6").FormulaR1C1 = "=RC[-1]*(1950-1920)"
    
    Range("K4:L4").HorizontalAlignment = xlCenter
    Range("K4:L4").Merge
    Range("K4:L4").FormulaR1C1 = "1950-" & Temp_last_year
    Range("K5").FormulaR1C1 = Chr(176) & "C/year"
    Range("L5").FormulaR1C1 = "Total " & ChrW(8710) & "T (" & Chr(176) & "C)"
    Range("L6").FormulaR1C1 = "=RC[-1]*(" & Temp_last_year & "-1950)"
    
    Range("B4:L10").Select
    With Selection
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
    End With

    Range("A14:F14").HorizontalAlignment = xlCenter
    Range("A14:F14").Merge
    Range("A14:F14").FormulaR1C1 = "Total Precipitation Change"
    
    Range("B16:C16").HorizontalAlignment = xlCenter
    Range("B16:C16").Merge
    Range("B16:C16").FormulaR1C1 = Prcp_first_year & "-" & Prcp_last_year
    
    Range("A17").FormulaR1C1 = "All Seasons"
    Range("B17").FormulaR1C1 = "mm/year"
    Range("C17").FormulaR1C1 = "Total " & ChrW(8710) & " Precipitation (mm)"
    Range("C18").FormulaR1C1 = "=RC[-1]*(" & Prcp_last_year & "-" & Prcp_first_year & ")"
    Range("F18").FormulaR1C1 = "=RC[-1]/RC[-3]"
    
    Range("E17").FormulaR1C1 = "Average Total Precipitation (mm)"
    Range("F17").FormulaR1C1 = "% Change Total Precipitation"
    
    Range("B16:F19").Select
    With Selection
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
    End With
     
    For i = 1 To 12
        Columns(Col_lett(i) & ":" & Col_lett(i)).EntireColumn.AutoFit
    Next i
    
End Sub
