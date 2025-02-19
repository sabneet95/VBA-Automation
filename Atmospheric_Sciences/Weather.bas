Option Explicit

'*******************************************************************************
' Module:       Weather Data Processing Macro
' Author:       Sabneet Bains
' Description:  Processes imported weather data (temperature and precipitation)
'               for a given station. It locates key header values, inserts new
'               calculation columns, fills formulas for seasonal and annual averages,
'               builds multiple charts (with trendlines), and arranges summary tables.
'
' Usage:        Call the Weather subroutine from the VBA editor or via an Excel
'               button. Ensure that the worksheet structure conforms to the assumed
'               layout (headers like STA_NAME, COUNTRY, STA_ID, TYPE, Year, JAN, DEC, etc.).
'
' Requirements: Microsoft Excel 2016 or later, VBA 7 or higher.
'
' License:      MIT License
'*******************************************************************************

'------------------------------------------------------------------------------
' Helper Function: Col_lett
' Returns the Excel column letter for a given column number.
'------------------------------------------------------------------------------
Function Col_lett(ByVal ColumnNumber As Integer) As String
    ' Remove row number and "$" signs from the cell address in row 1.
    Col_lett = Replace(Replace(Cells(1, ColumnNumber).Address, "1", ""), "$", "")
End Function

'------------------------------------------------------------------------------
' Main Subroutine: Weather
' Processes the imported weather data, creates formulas for temperature and 
' precipitation calculations, builds charts, and creates summary tables.
'------------------------------------------------------------------------------
Sub Weather()
    Dim city_name As String, country_name As String
    Dim STA_ID_row As Long, temp_last_row As Long, prcp_last_row As Long
    Dim type_column As Long, year_column As Long, january_column As Long, december_column As Long
    Dim Temp_first_year As Variant, Temp_last_year As Variant
    Dim Prcp_first_year As Variant, Prcp_last_year As Variant
    Dim i As Long, Chart_index As Long
    Dim y_1920 As Long, y_1950 As Long
    Dim Width As Single, Height As Single, NumWide As Long
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '==============================
    ' Locate Station Information
    '==============================
    Range("A1").Select
    city_name = Cells.Find(What:="*STA_NAME", LookIn:=xlFormulas, LookAt:=xlPart, _
                  SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Offset(1, 0).Value
    country_name = Cells.Find(What:="*COUNTRY", LookIn:=xlFormulas, LookAt:=xlPart, _
                     SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Offset(1, 0).Value

    '===============================================
    ' TEMPERATURE DATA PROCESSING
    '===============================================
    ' Locate temperature data start using header "*STA_ID"
    Cells.Find(What:="*STA_ID", LookIn:=xlFormulas, LookAt:=xlPart, _
               SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Activate
    STA_ID_row = ActiveCell.Row
    temp_last_row = ActiveCell.End(xlDown).Row

    ' Find key columns for temperature calculations
    type_column = Cells.Find(What:="*TYPE", LookIn:=xlFormulas, LookAt:=xlPart, _
                   SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Column
    year_column = Cells.Find(What:="*Year", LookIn:=xlFormulas, LookAt:=xlPart, _
                   SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Column
    january_column = Cells.Find(What:="*JAN", LookIn:=xlFormulas, LookAt:=xlPart, _
                     SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Column
    december_column = Cells.Find(What:="*DEC", LookIn:=xlFormulas, LookAt:=xlPart, _
                      SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Column

    '-----------------------------------------
    ' Insert Calculation Columns for Temperature
    '-----------------------------------------
    For i = 1 To 4
        Columns(Col_lett(december_column + 1)).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next i

    '---- Annual Average Temperature ----
    With Columns(Col_lett(december_column + 1))
        .Font.Bold = True
    End With
    With Range(Col_lett(december_column + 1) & STA_ID_row)
        .FormulaR1C1 = "AVERAGE ANNUAL TEMP"
    End With
    With Range(Col_lett(december_column + 1) & (STA_ID_row + 1))
        .FormulaR1C1 = "=AVERAGE(RC[" & -(december_column + 1 - january_column) & "]:RC[-1])"
        .AutoFill Destination:=Range(Col_lett(december_column + 1) & (STA_ID_row + 1) & ":" & _
                                      Col_lett(december_column + 1) & temp_last_row), Type:=xlFillDefault
    End With

    '---- Winter Average Temperature ----
    With Columns(Col_lett(december_column + 2))
        .Font.Bold = True
    End With
    With Range(Col_lett(december_column + 2) & STA_ID_row)
        .FormulaR1C1 = "AVERAGE WINTER TEMP"
    End With
    With Range(Col_lett(december_column + 2) & (STA_ID_row + 1))
        .FormulaR1C1 = "=AVERAGE(RC[-2],RC[" & -(december_column + 2 - january_column) & "]:RC[" & -(december_column + 1 - january_column) & "])"
        .AutoFill Destination:=Range(Col_lett(december_column + 2) & (STA_ID_row + 1) & ":" & _
                                      Col_lett(december_column + 2) & temp_last_row), Type:=xlFillDefault
    End With

    '---- Spring Average Temperature ----
    With Columns(Col_lett(december_column + 3))
        .Font.Bold = True
    End With
    With Range(Col_lett(december_column + 3) & STA_ID_row)
        .FormulaR1C1 = "AVERAGE SPRING TEMP"
    End With
    With Range(Col_lett(december_column + 3) & (STA_ID_row + 1))
        .FormulaR1C1 = "=AVERAGE(RC[" & -(december_column + 1 - january_column) & "]:RC[" & -(december_column - (january_column + 1)) & "])"
        .AutoFill Destination:=Range(Col_lett(december_column + 3) & (STA_ID_row + 1) & ":" & _
                                      Col_lett(december_column + 3) & temp_last_row), Type:=xlFillDefault
    End With

    '---- Summer Average Temperature ----
    With Columns(Col_lett(december_column + 4))
        .Font.Bold = True
    End With
    With Range(Col_lett(december_column + 4) & STA_ID_row)
        .FormulaR1C1 = "AVERAGE SUMMER TEMP"
    End With
    With Range(Col_lett(december_column + 4) & (STA_ID_row + 1))
        .FormulaR1C1 = "=AVERAGE(RC[" & -(december_column - (january_column + 1)) & "]:RC[" & -(december_column - (january_column + 3)) & "])"
        .AutoFill Destination:=Range(Col_lett(december_column + 4) & (STA_ID_row + 1) & ":" & _
                                      Col_lett(december_column + 4) & temp_last_row), Type:=xlFillDefault
    End With

    '---- Fall Average Temperature ----
    With Columns(Col_lett(december_column + 5))
        .Font.Bold = True
    End With
    With Range(Col_lett(december_column + 5) & STA_ID_row)
        .FormulaR1C1 = "AVERAGE FALL TEMP"
    End With
    With Range(Col_lett(december_column + 5) & (STA_ID_row + 1))
        .FormulaR1C1 = "=AVERAGE(RC[" & -(december_column - (january_column + 3)) & "]:RC[" & -(december_column - (january_column + 5)) & "])"
        .AutoFill Destination:=Range(Col_lett(december_column + 5) & (STA_ID_row + 1) & ":" & _
                                      Col_lett(december_column + 5) & temp_last_row), Type:=xlFillDefault
    End With

    '---- Bottom Row (Overall Annual Temperature Average) ----
    With Rows(temp_last_row + 1)
        .Font.Bold = True
    End With
    With Range(Col_lett(type_column) & (temp_last_row + 1))
        .FormulaR1C1 = "AVERAGE"
    End With
    With Range(Col_lett(january_column) & (temp_last_row + 1))
        .FormulaR1C1 = "=AVERAGE(R[" & -(temp_last_row - STA_ID_row) & "]C:R[-1]C)"
        .AutoFill Destination:=Range(Col_lett(january_column) & (temp_last_row + 1) & ":" & _
                                      Col_lett(december_column) & (temp_last_row + 1)), Type:=xlFillDefault
    End With

    ' Record first and last year from temperature data for chart titles
    Temp_first_year = Range(Col_lett(year_column) & (STA_ID_row + 1)).Value
    Temp_last_year = Range(Col_lett(year_column) & temp_last_row).Value

    '===============================================
    ' Build Temperature Charts
    '===============================================
    '--- Chart 1: Average Annual Temperature ---
    With ActiveSheet
        .Range(Col_lett(december_column + 1) & (STA_ID_row + 1) & ":" & _
               Col_lett(december_column + 1) & temp_last_row & "," & _
               Col_lett(year_column) & (STA_ID_row + 1) & ":" & _
               Col_lett(year_column) & temp_last_row).Select
    End With
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    With ActiveChart
        .FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(year_column) & "$" & temp_last_row
        .FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 1) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(december_column + 1) & "$" & temp_last_row
        .ChartTitle.Text = city_name & " " & country_name & " Average Annual Temp " & Temp_first_year & "-" & Temp_last_year & " Figure 1"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time (years)"
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue).TickLabels.Font.Size = 8
        With .FullSeriesCollection(1)
            .MarkerStyle = 8
            .MarkerSize = 4
            .Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
            .Trendlines.Add
            .Trendlines(1).DisplayEquation = True
            .Trendlines(1).DisplayRSquared = True
        End With
    End With

    '--- Chart 2: Average Annual Temp (up to 1920) ---
    Range(Col_lett(year_column) & STA_ID_row).Select
    y_1920 = Cells.Find(What:="1920", LookIn:=xlFormulas, LookAt:=xlPart, _
                         SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Row
    With ActiveSheet
        .Range(Col_lett(december_column + 1) & (STA_ID_row + 1) & ":" & _
               Col_lett(december_column + 1) & y_1920 & "," & _
               Col_lett(year_column) & (STA_ID_row + 1) & ":" & _
               Col_lett(year_column) & y_1920).Select
    End With
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    With ActiveChart
        .FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(year_column) & "$" & y_1920
        .FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 1) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(december_column + 1) & "$" & y_1920
        Range("Sheet3!E6").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 1) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 1) & y_1920 & ",Sheet1!" & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & y_1920 & ")"
        .ChartTitle.Text = city_name & " " & country_name & " Average Annual Temp " & Range(Col_lett(year_column) & (STA_ID_row + 1)).Value & "-1920 Figure 2"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time (years)"
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        With .FullSeriesCollection(1)
            .MarkerStyle = 8
            .MarkerSize = 4
            .Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
            .Trendlines.Add
            .Trendlines(1).DisplayEquation = True
            .Trendlines(1).DisplayRSquared = True
        End With
    End With

    '--- Chart 3: Average Annual Temp (1920-1950) ---
    Range(Col_lett(year_column) & STA_ID_row).Select
    y_1950 = Cells.Find(What:="1950", LookIn:=xlFormulas, LookAt:=xlPart, _
                         SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Row
    With ActiveSheet
        .Range(Col_lett(december_column + 1) & y_1920 & ":" & _
               Col_lett(december_column + 1) & y_1950 & "," & _
               Col_lett(year_column) & y_1920 & ":" & _
               Col_lett(year_column) & y_1950).Select
    End With
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    With ActiveChart
        .FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & y_1920 & ":" & "$" & Col_lett(year_column) & "$" & y_1950
        .FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 1) & "$" & y_1920 & ":" & "$" & Col_lett(december_column + 1) & "$" & y_1950
        Range("Sheet3!H6").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 1) & y_1920 & ":" & Col_lett(december_column + 1) & y_1950 & ",Sheet1!" & Col_lett(year_column) & y_1920 & ":" & Col_lett(year_column) & y_1950 & ")"
        .ChartTitle.Text = city_name & " " & country_name & " Average Annual Temp 1920-1950 Figure 3"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time (years)"
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        With .FullSeriesCollection(1)
            .MarkerStyle = 8
            .MarkerSize = 4
            .Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
            .Trendlines.Add
            .Trendlines(1).DisplayEquation = True
            .Trendlines(1).DisplayRSquared = True
        End With
    End With

    '--- Chart 4: Average Annual Temp (1950 to Last Year) ---
    With ActiveSheet
        .Range(Col_lett(december_column + 1) & y_1950 & ":" & _
               Col_lett(december_column + 1) & temp_last_row & "," & _
               Col_lett(year_column) & y_1950 & ":" & _
               Col_lett(year_column) & temp_last_row).Select
    End With
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    With ActiveChart
        .FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & y_1950 & ":" & "$" & Col_lett(year_column) & "$" & temp_last_row
        .FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 1) & "$" & y_1950 & ":" & "$" & Col_lett(december_column + 1) & "$" & temp_last_row
        Range("Sheet3!K6").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 1) & y_1950 & ":" & Col_lett(december_column + 1) & temp_last_row & ",Sheet1!" & Col_lett(year_column) & y_1950 & ":" & Col_lett(year_column) & temp_last_row & ")"
        .ChartTitle.Text = city_name & " " & country_name & " Average Annual Temp 1950-" & Range(Col_lett(year_column) & temp_last_row).Value & " Figure 4"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time (years)"
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        With .FullSeriesCollection(1)
            .MarkerStyle = 8
            .MarkerSize = 4
            .Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
            .Trendlines.Add
            .Trendlines(1).DisplayEquation = True
            .Trendlines(1).DisplayRSquared = True
        End With
    End With

    '--- Chart 5: Average Winter (DJF) Temperature ---
    With ActiveSheet
        .Range(Col_lett(december_column + 2) & (STA_ID_row + 1) & ":" & _
               Col_lett(december_column + 2) & temp_last_row & "," & _
               Col_lett(year_column) & (STA_ID_row + 1) & ":" & _
               Col_lett(year_column) & temp_last_row).Select
    End With
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    With ActiveChart
        .FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(year_column) & "$" & temp_last_row
        .FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 2) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(december_column + 2) & "$" & temp_last_row
        Range("Sheet3!B7").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 2) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 2) & temp_last_row & ",Sheet1!" & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & temp_last_row & ")"
        .ChartTitle.Text = city_name & " " & country_name & " Average Winter(DJF) Temp " & Range(Col_lett(year_column) & (STA_ID_row + 1)).Value & "-" & Range(Col_lett(year_column) & temp_last_row).Value & " Figure 5"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time (years)"
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        With .FullSeriesCollection(1)
            .MarkerStyle = 8
            .MarkerSize = 4
            .Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
            .Trendlines.Add
            .Trendlines(1).DisplayEquation = True
            .Trendlines(1).DisplayRSquared = True
        End With
    End With

    '--- Chart 6: Average Spring (MAM) Temperature ---
    With ActiveSheet
        .Range(Col_lett(december_column + 3) & (STA_ID_row + 1) & ":" & _
               Col_lett(december_column + 3) & temp_last_row & "," & _
               Col_lett(year_column) & (STA_ID_row + 1) & ":" & _
               Col_lett(year_column) & temp_last_row).Select
    End With
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    With ActiveChart
        .FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(year_column) & "$" & temp_last_row
        .FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 3) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(december_column + 3) & "$" & temp_last_row
        Range("Sheet3!B8").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 3) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 3) & temp_last_row & ",Sheet1!" & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & temp_last_row & ")"
        .ChartTitle.Text = city_name & " " & country_name & " Average Spring(MAM) Temp " & Range(Col_lett(year_column) & (STA_ID_row + 1)).Value & "-" & Range(Col_lett(year_column) & temp_last_row).Value & " Figure 6"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time (years)"
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        With .FullSeriesCollection(1)
            .MarkerStyle = 8
            .MarkerSize = 4
            .Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
            .Trendlines.Add
            .Trendlines(1).DisplayEquation = True
            .Trendlines(1).DisplayRSquared = True
        End With
    End With

    '--- Chart 7: Average Summer (JJA) Temperature ---
    With ActiveSheet
        .Range(Col_lett(december_column + 4) & (STA_ID_row + 1) & ":" & _
               Col_lett(december_column + 4) & temp_last_row & "," & _
               Col_lett(year_column) & (STA_ID_row + 1) & ":" & _
               Col_lett(year_column) & temp_last_row).Select
    End With
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    With ActiveChart
        .FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(year_column) & "$" & temp_last_row
        .FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 4) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(december_column + 4) & "$" & temp_last_row
        Range("Sheet3!B9").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 4) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 4) & temp_last_row & ",Sheet1!" & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & temp_last_row & ")"
        .ChartTitle.Text = city_name & " " & country_name & " Average Summer(JJA) Temp " & Range(Col_lett(year_column) & (STA_ID_row + 1)).Value & "-" & Range(Col_lett(year_column) & temp_last_row).Value & " Figure 7"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time (years)"
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        With .FullSeriesCollection(1)
            .MarkerStyle = 8
            .MarkerSize = 4
            .Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
            .Trendlines.Add
            .Trendlines(1).DisplayEquation = True
            .Trendlines(1).DisplayRSquared = True
        End With
    End With

    '--- Chart 8: Average Fall (SON) Temperature ---
    With ActiveSheet
        .Range(Col_lett(december_column + 5) & (STA_ID_row + 1) & ":" & _
               Col_lett(december_column + 5) & temp_last_row & "," & _
               Col_lett(year_column) & (STA_ID_row + 1) & ":" & _
               Col_lett(year_column) & temp_last_row).Select
    End With
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    With ActiveChart
        .FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(year_column) & "$" & temp_last_row
        .FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 5) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(december_column + 5) & "$" & temp_last_row
        Range("Sheet3!B10").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 5) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 5) & temp_last_row & ",Sheet1!" & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & temp_last_row & ")"
        .ChartTitle.Text = city_name & " " & country_name & " Average Fall(SON) Temp " & Range(Col_lett(year_column) & (STA_ID_row + 1)).Value & "-" & Range(Col_lett(year_column) & temp_last_row).Value & " Figure 8"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time (years)"
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Temperature (" & Chr(176) & "C)"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        With .FullSeriesCollection(1)
            .MarkerStyle = 8
            .MarkerSize = 4
            .Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
            .Trendlines.Add
            .Trendlines(1).DisplayEquation = True
            .Trendlines(1).DisplayRSquared = True
        End With
    End With

    '===============================================
    ' PRECIPITATION DATA PROCESSING
    '===============================================
    ' Re-locate STA_ID row (for precipitation data) and determine last row
    Range("A" & temp_last_row).Select
    Cells.Find(What:="*STA_ID", LookIn:=xlFormulas, LookAt:=xlPart, _
               SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Activate
    STA_ID_row = ActiveCell.Row
    prcp_last_row = ActiveCell.End(xlDown).Row

    ' Find key columns for precipitation calculations
    type_column = Cells.Find(What:="*TYPE", LookIn:=xlFormulas, LookAt:=xlPart, _
                   SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Column
    january_column = Cells.Find(What:="*JAN", LookIn:=xlFormulas, LookAt:=xlPart, _
                      SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Column
    december_column = Cells.Find(What:="*DEC", LookIn:=xlFormulas, LookAt:=xlPart, _
                       SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Column

    '---- Total Precipitation ----
    With Columns(Col_lett(december_column + 1))
        .Font.Bold = True
    End With
    With Range(Col_lett(december_column + 1) & STA_ID_row)
        .FormulaR1C1 = "TOTAL PRECIPITATION"
    End With
    With Range(Col_lett(december_column + 1) & (STA_ID_row + 1))
        .FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
        .AutoFill Destination:=Range(Col_lett(december_column + 1) & (STA_ID_row + 1) & ":" & _
                                      Col_lett(december_column + 1) & prcp_last_row), Type:=xlFillDefault
    End With

    '---- Winter Total Precipitation ----
    With Range(Col_lett(december_column + 2) & STA_ID_row)
        .FormulaR1C1 = "TOTAL WINTER PRECIPITATION"
    End With
    With Range(Col_lett(december_column + 2) & (STA_ID_row + 1))
        .FormulaR1C1 = "=SUM(RC[-2],RC[" & -(december_column + 2 - january_column) & "]:RC[" & -(december_column + 1 - january_column) & "])"
        .AutoFill Destination:=Range(Col_lett(december_column + 2) & (STA_ID_row + 1) & ":" & _
                                      Col_lett(december_column + 2) & prcp_last_row), Type:=xlFillDefault
    End With

    '---- Spring Total Precipitation ----
    With Range(Col_lett(december_column + 3) & STA_ID_row)
        .FormulaR1C1 = "TOTAL SPRING PRECIPITATION"
    End With
    With Range(Col_lett(december_column + 3) & (STA_ID_row + 1))
        .FormulaR1C1 = "=SUM(RC[" & -(december_column + 1 - january_column) & "]:RC[" & -(december_column - (january_column + 1)) & "])"
        .AutoFill Destination:=Range(Col_lett(december_column + 3) & (STA_ID_row + 1) & ":" & _
                                      Col_lett(december_column + 3) & prcp_last_row), Type:=xlFillDefault
    End With

    '---- Summer Total Precipitation ----
    With Range(Col_lett(december_column + 4) & STA_ID_row)
        .FormulaR1C1 = "TOTAL SUMMER PRECIPITATION"
    End With
    With Range(Col_lett(december_column + 4) & (STA_ID_row + 1))
        .FormulaR1C1 = "=SUM(RC[" & -(december_column - (january_column + 1)) & "]:RC[" & -(december_column - (january_column + 3)) & "])"
        .AutoFill Destination:=Range(Col_lett(december_column + 4) & (STA_ID_row + 1) & ":" & _
                                      Col_lett(december_column + 4) & prcp_last_row), Type:=xlFillDefault
    End With

    '---- Fall Total Precipitation ----
    With Range(Col_lett(december_column + 5) & STA_ID_row)
        .FormulaR1C1 = "TOTAL FALL PRECIPITATION"
    End With
    With Range(Col_lett(december_column + 5) & (STA_ID_row + 1))
        .FormulaR1C1 = "=SUM(RC[" & -(december_column - (january_column + 3)) & "]:RC[" & -(december_column - (january_column + 5)) & "])"
        .AutoFill Destination:=Range(Col_lett(december_column + 5) & (STA_ID_row + 1) & ":" & _
                                      Col_lett(december_column + 5) & prcp_last_row), Type:=xlFillDefault
    End With

    '---- Bottom Row for Precipitation Average ----
    With Rows(prcp_last_row + 1)
        .Font.Bold = True
    End With
    With Range(Col_lett(type_column) & (prcp_last_row + 1))
        .FormulaR1C1 = "AVERAGE"
    End With
    With Range(Col_lett(january_column) & (prcp_last_row + 1))
        .FormulaR1C1 = "=AVERAGE(R[" & -(prcp_last_row - STA_ID_row) & "]C:R[-1]C)"
        .AutoFill Destination:=Range(Col_lett(january_column) & (prcp_last_row + 1) & ":" & _
                                      Col_lett(december_column + 1) & (prcp_last_row + 1)), Type:=xlFillDefault
    End With
    Range("Sheet3!E18").Value = Range(Col_lett(december_column + 1) & (prcp_last_row + 1)).Value

    Prcp_first_year = Range(Col_lett(year_column) & (STA_ID_row + 1)).Value
    Prcp_last_year = Range(Col_lett(year_column) & prcp_last_row).Value

    '--- Chart 9: Total Annual Precipitation ---
    With ActiveSheet
        .Range(Col_lett(december_column + 1) & (STA_ID_row + 1) & ":" & _
               Col_lett(december_column + 1) & prcp_last_row & "," & _
               Col_lett(year_column) & (STA_ID_row + 1) & ":" & _
               Col_lett(year_column) & prcp_last_row).Select
    End With
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    With ActiveChart
        .FullSeriesCollection(1).XValues = "=Sheet1!" & "$" & Col_lett(year_column) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(year_column) & "$" & prcp_last_row
        .FullSeriesCollection(1).Values = "=Sheet1!" & "$" & Col_lett(december_column + 1) & "$" & (STA_ID_row + 1) & ":" & "$" & Col_lett(december_column + 1) & "$" & prcp_last_row
        Range("Sheet3!B18").Value = "=LINEST(Sheet1!" & Col_lett(december_column + 1) & (STA_ID_row + 1) & ":" & Col_lett(december_column + 1) & prcp_last_row & ",Sheet1!" & Col_lett(year_column) & (STA_ID_row + 1) & ":" & Col_lett(year_column) & prcp_last_row & ")"
        .ChartTitle.Text = city_name & " " & country_name & " Total Annual Precipitation " & Range(Col_lett(year_column) & (STA_ID_row + 1)).Value & "-" & Range(Col_lett(year_column) & prcp_last_row).Value & " Figure 9"
        .ChartTitle.Font.Size = 10
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time (years)"
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).MajorTickMark = xlOutside
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Total Precipitation (mm)"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlOutside
        With .FullSeriesCollection(1)
            .MarkerStyle = 8
            .MarkerSize = 4
            .Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
            .Trendlines.Add
            .Trendlines(1).DisplayEquation = True
            .Trendlines(1).DisplayRSquared = True
        End With
    End With

    '===============================================
    ' FINAL FORMATTING, CHART ARRANGEMENT & SUMMARY TABLES
    '===============================================
    ' Auto-fit the first 12 columns for readability
    For i = 1 To 12
        Columns(Col_lett(i)).EntireColumn.AutoFit
    Next i

    ' Arrange charts by resizing and repositioning them
    Dim Chart_count As Long
    Chart_count = ActiveSheet.ChartObjects.Count
    Width = 480: Height = 280
    NumWide = (Chart_count / 3)
    For Chart_index = 1 To Chart_count
        With ActiveSheet.ChartObjects(Chart_index)
            .Width = Width
            .Height = Height
            .Left = ((Chart_index - 1) Mod NumWide) * Width
            .Top = Int((Chart_index - 1) / NumWide) * Height
        End With
    Next Chart_index
    For Chart_index = 1 To Chart_count / 3
        ActiveSheet.ChartObjects(Chart_index).Top = 620
    Next Chart_index
    For Chart_index = Chart_count / 3 + 1 To ((2 * Chart_count) / 3)
        ActiveSheet.ChartObjects(Chart_index).Top = 900
    Next Chart_index
    For Chart_index = ((2 * Chart_count) / 3) + 1 To Chart_count
        ActiveSheet.ChartObjects(Chart_index).Top = 1180
    Next Chart_index

    ' Example of moving charts to a separate "Graphs" sheet:
    Sheets(ActiveSheet.Name).Name = "Main_Data"
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes.Range(Array("Chart 1", "Chart 2", "Chart 3", "Chart 4", _
        "Chart 5", "Chart 6", "Chart 7", "Chart 8", "Chart 9")).Select
    Selection.Cut
    Sheets("Sheet2").Select
    Range("A3").Select
    ActiveSheet.Paste
    Sheets(ActiveSheet.Name).Name = "Graphs"
    
    ' Build Summary Tables in "Summary_Tables" sheet:
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
        Columns(Col_lett(i)).EntireColumn.AutoFit
    Next i

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Weather Macro Error"
    Resume Cleanup
End Sub
