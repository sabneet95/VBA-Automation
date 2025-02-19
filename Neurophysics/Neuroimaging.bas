Option Explicit

'*******************************************************************************
' Module:       Neuro Data Processing and Graph Generation
' Author:       Sabneet Bains
' Description:  Processes neuroimaging data by upgrading non-XML files,
'               rearranging frame sequences, converting ROI names and frame 
'               numbers, creating SUV and SUVr calculations, and generating
'               various graphs and multiplots. Also creates time columns,
'               applies formatting, and builds summary data.
'
' Usage:        Run the Neuro subroutine from the VBA editor or via an Excel
'               button. Make sure the workbook and sheet structure match the 
'               assumed layout.
'
' Requirements: Microsoft Excel 2016 or later, VBA 7 or higher.
'
' License:      MIT License
'*******************************************************************************
Public aborting_mechanism As Integer

'------------------------------------------------------------------------------
' Function: onlyDigits
' Returns a string containing only the numeric digits found in the input.
'------------------------------------------------------------------------------
Function onlyDigits(s As String) As String
    Dim retval As String
    Dim i As Integer
    retval = ""
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval & Mid(s, i, 1)
        End If
    Next i
    onlyDigits = retval
End Function

'------------------------------------------------------------------------------
' Function: Col_lett
' Returns the column letter (e.g., "A", "B", etc.) for a given column number.
'------------------------------------------------------------------------------
Function Col_lett(ByVal ColumnNumber As Integer) As String
    Col_lett = Replace(Replace(Cells(1, ColumnNumber).Address, "1", ""), "$", "")
End Function

'------------------------------------------------------------------------------
' Subroutine: Neuro
' Main routine that upgrades the file, rearranges frame sequences, processes ROI 
' and SUV/SUVr calculations, and creates graphs and summary tables.
'------------------------------------------------------------------------------
Sub Neuro()
    Dim Current_path As String, Extension As String, Upgraded_File As String
    Dim Frames As VbMsgBoxResult
    Dim column_total As Long, row_total As Long, Initial_delay_divide As Long
    Dim Subject_value As String, Subject_frame_number As String, Subject_study As String
    Dim Subject_length As Long, Subject_underscore As Long
    Dim C As Range, D As Range, Digit_value As Variant
    Dim ROIs As Variant, ROIs_names As Variant
    Dim ROIs_length As Long, Sub_ROIs_length As Long
    Dim r As Long, rr As Long
    Dim AddressArr(10) As Variant ' to store cell addresses
    Dim column_addition As String
    Dim FramesAnswer As VbMsgBoxResult
    Dim Decay_row As Long
    Dim Weight As Double, Dose As Double
    Dim Weight_in_grams As Double, Dose_in_Bq As Double
    Dim Weight_row As Long, Dose_row As Long
    Dim SUV_formula As String, Decay_SUV_formula As String
    Dim SUVr_Formula As String, MAX_formula As String, MIN_formula As String
    Dim MAX_MAX_formula As String, MIN_MIN_formula As String
    Dim SUV_max As Double, SUV_min As Double, SUVr_max As Double, SUVr_min As Double
    Dim SUVr2_max As Double, SUVr2_min As Double
    Dim Sheet_Name As String
    Dim Width As Single, Height As Single, NumWide As Long
    Dim Chart_index As Long, Chart_count As Long
    Dim Col_array(3, 8) As Variant
    Dim Multi_range As String, Multiplot As String
    Dim vAxis As Variant
    Dim column_total_after_SUV As Long, column_total_after_SUVr As Long, column_total_after_SUVr_and_Time As Long
    Dim P As Long, q As Long, r2 As Long, s As Long, t As Long, u As Long, pIndex As Long
    Dim Multiplot_index As Long
    Dim SUV_Range As String
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '==============================
    ' File Upgrade (if needed)
    '==============================
    Current_path = Application.ActiveWorkbook.FullName
    Extension = Right(Current_path, 1)
    If Extension <> "x" Then
        Upgraded_File = Current_path & "x"
        ActiveWorkbook.SaveAs Filename:=Upgraded_File, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        Workbooks.Open Filename:=Upgraded_File
    End If

    '==============================
    ' Delete Mean Frame if present
    '==============================
    Range("A1").Select
    If InStr(Range("A2").Value, "mean") > 0 Then
        Rows("2:2").Delete Shift:=xlUp
    End If

    '==============================
    ' Frame Sequence Fix
    '==============================
    FramesAnswer = MsgBox("Would you like to rearrange the frames ascendingly by time?", vbYesNo + vbQuestion, "Frame Order")
    If FramesAnswer = vbYes Then
        ' Instead of hard-coded row moves (commented out), perform a sort.
        column_total = Range("A1").End(xlToRight).Offset(0, 4).Column
        row_total = Range("A1").End(xlDown).Offset(0, 1).Row
        
        With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("A2"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            .SetRange Range("A2:" & Col_lett(column_total) & row_total)
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        With Range("A1:A" & row_total)
            .Replace What:=".img", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
            .Replace What:="wrrxx", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
            Set C = .Find(What:="*_d*_f", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
            Do While Not C Is Nothing
                Subject_value = C.Value
                Subject_length = Len(Subject_value)
                Subject_underscore = InStr(UCase(Subject_value), "_D")
                Subject_frame_number = Right(Subject_value, Subject_length - Subject_underscore)
                C.Value = onlyDigits(Subject_frame_number)
                Set C = .FindNext(C)
            Loop
            
            Set D = .Find(What:=Left(.Cells(1, 1).Value, InStr(.Cells(1, 1).Value, "_f") - 1) & "_f", LookIn:=xlValues)
            Do While Not D Is Nothing
                Subject_value = D.Value
                Subject_length = Len(Subject_value)
                Subject_underscore = InStr(UCase(Subject_value), "_F")
                Subject_frame_number = Right(Subject_value, Subject_length - Subject_underscore)
                D.Value = onlyDigits(Subject_frame_number)
                Set D = .FindNext(D)
            Loop
        End With
        
        Initial_delay_divide = ActiveCell.Row
        With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("A2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A2:" & Col_lett(column_total) & Initial_delay_divide - 1)
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("A" & Initial_delay_divide), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A" & Initial_delay_divide & ":" & Col_lett(column_total) & row_total)
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        Dim DIndex As Long
        For DIndex = 2 To Initial_delay_divide - 1
            Digit_value = Range("A" & DIndex).Value
            Range("A" & DIndex).Value = Left(Subject_value, InStr(Subject_value, "_f") - 1) & "_Frame" & Digit_value
        Next DIndex
        
        Dim CIndex As Long
        For CIndex = Initial_delay_divide To row_total
            Digit_value = Range("A" & CIndex).Value
            Range("A" & CIndex).Value = Left(Subject_value, InStr(Subject_value, "_f") - 1) & "_Delayed_Frame" & Digit_value
        Next CIndex
    End If

    '==============================
    ' ROIs Creation
    '==============================
    ROIs = Array( _
        Array("Lingual_L_AAL.nii", "Occipital_Sup_L_AAL.nii", "Occipital_Mid_L_AAL.nii", "Occipital_Inf_L_AAL.nii", "Cuneus_L_AAL.nii", "Calcarine_L_AAL.nii"), _
        Array("Lingual_R_AAL.nii", "Occipital_Sup_R_AAL.nii", "Occipital_Mid_R_AAL.nii", "Occipital_Inf_R_AAL.nii", "Cuneus_R_AAL.nii", "Calcarine_R_AAL.nii"), _
        Array("Angular_L_AAL.nii", "SupraMarginal_L_AAL.nii", "Parietal_Sup_L_AAL.nii", "Parietal_Inf_L_AAL.nii", "Precuneus_L_AAL.nii"), _
        Array("Angular_R_AAL.nii", "SupraMarginal_R_AAL.nii", "Parietal_Sup_R_AAL.nii", "Parietal_Inf_R_AAL.nii", "Precuneus_R_AAL.nii"), _
        Array("Temporal_Pole_Mid_L_AAL.nii", "Temporal_Sup_L_AAL.nii", "Temporal_Pole_Mid_L_AAL.nii", "Temporal_Mid_L_AAL.nii", "Temporal_Inf_L_AAL.nii"), _
        Array("Temporal_Pole_Mid_R_AAL.nii", "Temporal_Sup_R_AAL.nii", "Temporal_Pole_Mid_R_AAL.nii", "Temporal_Mid_R_AAL.nii", "Temporal_Inf_R_AAL.nii"), _
        Array("Frontal_Sup_L_AAL.nii", "Frontal_Mid_L_AAL.nii", "Frontal_Inf_Oper_L_AAL.nii", "Frontal_Inf_Tri_L_AAL.nii", "Frontal_Sup_Medial_L_AAL.nii", "Supp_Motor_Area_L_AAL.nii"), _
        Array("Frontal_Sup_R_AAL.nii", "Frontal_Mid_R_AAL.nii", "Frontal_Inf_Oper_R_AAL.nii", "Frontal_Inf_Tri_R_AAL.nii", "Frontal_Sup_Medial_R_AAL.nii", "Supp_Motor_Area_R_AAL.nii"))
    ROIs_names = Array("Occipital_L", "Occipital_R", "Parietal_L", "Parietal_R", "Temporal_L", "Temporal_R", "Frontal_L", "Frontal_R")
    ROIs_length = UBound(ROIs)
    
    For r = 0 To ROIs_length
        Sub_ROIs_length = UBound(ROIs(r))
        For rr = 0 To Sub_ROIs_length
            ' Find the ROI header cell and store its column letter
            Range("C1").Select
            Cells.Find(What:=ROIs(r)(rr), LookIn:=xlFormulas, LookAt:=xlPart, _
                SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=True).Activate
            Col_array(r, rr) = Col_lett(Selection.Column)
        Next rr
        
        ' Build the average formula for this ROI and fill down
        With Range("B2:B" & row_total)
            ' Create a multi-range string from the stored column letters
            Multi_range = Col_array(r, 0) & "2:" & Col_array(r, 0) & row_total & "," & _
                          Col_array(r, 1) & "2:" & Col_array(r, 1) & row_total & "," & _
                          Col_array(r, 2) & "2:" & Col_array(r, 2) & row_total & "," & _
                          Col_array(r, 3) & "2:" & Col_array(r, 3) & row_total & "," & _
                          Col_array(r, 4) & "2:" & Col_array(r, 4) & row_total & "," & _
                          Col_array(r, 5) & "2:" & Col_array(r, 5) & row_total & "," & _
                          Col_array(r, 6) & "2:" & Col_array(r, 6) & row_total & "," & _
                          Col_array(r, 7) & "2:" & Col_array(r, 7) & row_total & "," & _
                          Col_array(r, 8) & "2:" & Col_array(r, 8) & row_total
        End With
        
        ' Set the ROI name in the header cell
        Range(Cells(1, ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column + 1), Cells(1, ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column + 1)).Value = ROIs_names(r)
        ' In the cell below, enter the average formula and autofill down
        With ActiveCell.Offset(1, 0)
            If Sub_ROIs_length = 4 Then
                .Value = "=AVERAGE(" & Col_array(r, 0) & "2," & Col_array(r, 1) & "2," & Col_array(r, 2) & "2," & Col_array(r, 3) & "2," & Col_array(r, 4) & "2)"
            ElseIf Sub_ROIs_length = 5 Then
                .Value = "=AVERAGE(" & Col_array(r, 0) & "2," & Col_array(r, 1) & "2," & Col_array(r, 2) & "2," & Col_array(r, 3) & "2," & Col_array(r, 4) & "2," & Col_array(r, 5) & "2)"
            End If
            .AutoFill Destination:=Range(.Address, Col_lett(ActiveCell.Column) & row_total), Type:=xlFillDefault
        End With
        
        If r = ROIs_length Then
            Range("A1").End(xlToRight).Offset(0, 1).Value = "Keep Blank!"
        End If
    Next r

    '==============================
    ' Rows & Columns Count & Formatting
    '==============================
    column_total = Range("A1").End(xlToRight).Column
    row_total = Range("A1").End(xlDown).Row
    
    Range("A1").Select
    Cells.Replace What:=".img", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
    Cells.Replace What:="wrrxx", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
    Columns(1).EntireColumn.AutoFit
    With Selection.Interior
        .Color = 6684876
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
    End With
    Selection.Font.Bold = True
    Range(Col_lett(2) & "1:" & Col_lett(column_total - 3) & "1").Select
    With Selection.Interior
        .Color = 13421619
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
    End With
    Selection.Font.Bold = True

    '==============================
    ' Time Columns Creation (Method 2 via UserForm)
    '==============================
    Columns(1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range(Col_lett(1) & "1").Value = "Time Intervals"
    Columns(2).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range(Col_lett(2) & "1").Value = "Start Time"
    UserForm1.Show
    ' If the flag set in the userform indicates to abort, then exit.
    If Flag > 0 Then Exit Sub
    Range("B1").Select
    Cells.Find(What:="90", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Activate
    Decay_row = ActiveCell.Row

    '==============================
    ' Weight and Dose Inputs
    '==============================
    Weight = InputBox("Now, please proceed to MIM and acquire the subject weight in Kg (by design Kg -> g):", "Subject Weight")
    Weight_in_grams = Weight * 1000#
    Dose = InputBox("And now, input the total dose registered in mCi (by design mCi -> Bq):", "Total Dose")
    Dose_in_Bq = Dose * 37000000#
    With Range("A1").End(xlDown).Offset(2, 0)
        .Value = "Patient Weight:"
    End With
    With Range("B1").End(xlDown).Offset(2, 0)
        .Value = "Total Dose:"
    End With
    With Range("A1").End(xlDown).Offset(3, 0)
        .Value = Weight_in_grams & " g"
        .ClearComments
        .AddComment "Please, keep the same format when modifying as it might accidentally break the rest of the functions! For your reference, copy the following default, if in vain: 10000 g"
        .Comment.Visible = True
        .Comment.Shape.ScaleWidth 1.5, msoFalse, msoScaleFromTopLeft
        .Comment.Shape.ScaleHeight 0.9, msoFalse, msoScaleFromTopLeft
        .Comment.Visible = False
    End With
    Weight_row = Range("A1").End(xlDown).Offset(3, 0).Row

    With Range("B1").End(xlDown).Offset(3, 0)
        .Value = Dose_in_Bq & " Bq"
        .ClearComments
        .AddComment "Same here, the default is: 11100000 Bq"
        .Comment.Visible = True
        .Comment.Shape.ScaleWidth 1, msoFalse, msoScaleFromTopLeft
        .Comment.Shape.ScaleHeight 0.9, msoFalse, msoScaleFromTopLeft
        .Comment.Visible = False
    End With
    Dose_row = Range("B1").End(xlDown).Offset(3, 0).Row

    '==============================
    ' SUV Column Creation and Calculations
    '==============================
    i = 6
    Do While i < column_total
        Columns(i).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range(Col_lett(i) & "1").Value = "=CONCATENATE(""SUV_"",RC[-2])"
        With Range(Col_lett(i) & "1").Interior
            .Color = 49407
        End With
        With Range(Col_lett(i) & "1").Font
            .ThemeColor = xlThemeColorDark1
        End With
        Range(Col_lett(i) & "1").Font.Bold = True
        SUV_formula = "=RC[-2]*(LEFT(R" & Weight_row & "C1,LEN(R" & Weight_row & "C1)-2))/(LEFT(R" & Dose_row & "C2,LEN(R" & Dose_row & "C2)-3))"
        Decay_SUV_formula = "=RC[-2]*(LEFT(R" & Weight_row & "C1,LEN(R" & Weight_row & "C1)-2))/((LEFT(R" & Dose_row & "C2,LEN(R" & Dose_row & "C2)-3))*EXP(-0.693*90/109.77))"
        Range(Col_lett(i) & "2").FormulaR1C1 = SUV_formula
        Range(Col_lett(i) & Decay_row).FormulaR1C1 = Decay_SUV_formula
        Range(Col_lett(i) & "2").AutoFill Destination:=Range(Col_lett(i) & "2:" & Col_lett(i) & Decay_row - 1), Type:=xlFillDefault
        Range(Col_lett(i) & Decay_row).AutoFill Destination:=Range(Col_lett(i) & Decay_row & ":" & Col_lett(i) & row_total), Type:=xlFillDefault
        i = i + 3
        column_total = column_total + 1
    Loop

    '==============================
    ' SUVr Column Creation and Calculations
    '==============================
    Range("A1").End(xlToRight).Offset(0, 3).Select
    column_total_after_SUV = ActiveCell.Column
    Range(Col_lett(1) & "1").Select ' reset view
    Dim j As Long
    j = 7
    Do While j < column_total_after_SUV
        Columns(j).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range(Col_lett(j) & "1").Value = "=CONCATENATE(""SUVr(blcere)_"",RC[-3])"
        With Range(Col_lett(j) & "1").Interior
            .Color = 49407
        End With
        With Range(Col_lett(j) & "1").Font
            .ThemeColor = xlThemeColorDark1
        End With
        Range(Col_lett(j) & "1").Font.Bold = True
        j = j + 4
        column_total_after_SUV = column_total_after_SUV + 1
    Loop
    
    j = 8
    Do While j < column_total_after_SUV
        Columns(j).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range(Col_lett(j) & "1").Value = "=CONCATENATE(""SUVr(Cerecrus)_"",RC[-4])"
        With Range(Col_lett(j) & "1").Interior
            .Color = 49407
        End With
        With Range(Col_lett(j) & "1").Font
            .ThemeColor = xlThemeColorDark1
        End With
        Range(Col_lett(j) & "1").Font.Bold = True
        j = j + 5
        column_total_after_SUV = column_total_after_SUV + 1
    Loop
    Range("A1").Select
    Cells.Replace What:=".nii", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False

    '==============================
    ' Graphs Creation - MAX/MIN Calculations and Chart Generation
    '==============================
    Range("A1").End(xlToRight).Offset(0, 2).Select
    column_total_after_SUVr_and_Time = ActiveCell.Column
    Range("A1").End(xlDown).Offset(0, 1).Select
    row_total = ActiveCell.Row

    Dim P As Long, q As Long, r2 As Long
    P = 6
    Do While P < column_total_after_SUVr_and_Time
        MAX_formula = "=MAX(R[-" & row_total + 3 & "]C:R[-1]C)"
        MIN_formula = "=MIN(R[-" & row_total + 4 & "]C:R[-2]C)"
        Range(Col_lett(P) & row_total + 5).Value = MAX_formula
        Range(Col_lett(P) & row_total + 6).Value = MIN_formula
        P = P + 5
    Loop

    Dim qIndex As Long
    qIndex = 7
    Do While qIndex < column_total_after_SUVr_and_Time
        MAX_formula = "=MAX(R[-" & row_total + 5 & "]C:R[-1]C)"
        MIN_formula = "=MIN(R[-" & row_total + 6 & "]C:R[-2]C)"
        Range(Col_lett(qIndex) & row_total + 7).Value = MAX_formula
        Range(Col_lett(qIndex) & row_total + 8).Value = MIN_formula
        qIndex = qIndex + 5
    Loop

    Dim rIndex As Long
    rIndex = 8
    Do While rIndex < column_total_after_SUVr_and_Time
        MAX_formula = "=MAX(R[-" & row_total + 7 & "]C:R[-1]C)"
        MIN_formula = "=MIN(R[-" & row_total + 8 & "]C:R[-2]C)"
        Range(Col_lett(rIndex) & row_total + 9).Value = MAX_formula
        Range(Col_lett(rIndex) & row_total + 10).Value = MIN_formula
        rIndex = rIndex + 5
    Loop

    MAX_MAX_formula = "=MAX(RC[-" & column_total - 8 & "]:RC[-1])"
    MIN_MIN_formula = "=MIN(RC[-" & column_total - 8 & "]:RC[-1])"
    With Range(Col_lett(column_total_after_SUVr_and_Time) & row_total)
        .Offset(5, -2).Value = MAX_MAX_formula
        SUV_max = WorksheetFunction.RoundUp(.Offset(5, -2).Value, 1)
        .Offset(6, -2).Value = MIN_MIN_formula
        SUV_min = WorksheetFunction.RoundDown(.Offset(6, -2).Value, 1)
        .Offset(7, -2).Value = MAX_MAX_formula
        SUVr_max = WorksheetFunction.RoundUp(.Offset(7, -2).Value, 1)
        .Offset(8, -2).Value = MIN_MIN_formula
        SUVr_min = WorksheetFunction.RoundDown(.Offset(8, -2).Value, 1)
        .Offset(9, -2).Value = MAX_MAX_formula
        SUVr2_max = WorksheetFunction.RoundUp(.Offset(9, -2).Value, 1)
        .Offset(10, -2).Value = MIN_MIN_formula
        SUVr2_min = WorksheetFunction.RoundDown(.Offset(10, -2).Value, 1)
    End With
    SUVr_max = WorksheetFunction.Max(SUVr_max, SUVr2_max)
    SUVr_min = WorksheetFunction.Min(SUVr_min, SUVr2_min)
    Rows(row_total + 5 & ":" & row_total + 10).Delete Shift:=xlUp
    Sheet_Name = ActiveSheet.Name

    '------------------------------
    ' Create Individual Graphs for SUV and SUVr
    '------------------------------
    s = 6
    Do While s < column_total_after_SUVr_and_Time
        ' Create chart for SUV columns
        Range("B2:B" & row_total).Select
        Range(Col_lett(s) & "2:" & Col_lett(s) & row_total).Select
        ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
        SUV_Range = Sheet_Name & "!$B$2:$B$" & row_total & "," & Sheet_Name & "!$" & Col_lett(s) & "$2:$" & Col_lett(s) & "$" & row_total
        With ActiveChart
            .SetSourceData Source:=Range(SUV_Range)
            .ChartTitle.Text = Range(Col_lett(s) & "1").Value
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time (mins)"
            .Axes(xlCategory).MinimumScale = 0
            .Axes(xlCategory).MaximumScale = 150
            .Axes(xlCategory).MajorUnit = 30
            .Axes(xlValue, xlPrimary).HasTitle = True
            If s < 8 Then
                .Axes(xlValue, xlPrimary).AxisTitle.Text = "SUV"
            Else
                .Axes(xlValue, xlPrimary).AxisTitle.Text = "SUVr"
            End If
        End With
        s = s + 5
    Loop
    
    ' Create additional graphs for SUVr using similar loops (using t and u)
    t = 7
    Do While t < column_total_after_SUVr_and_Time
        Range("B2:B" & row_total).Select
        Range(Col_lett(t) & "2:" & Col_lett(t) & row_total).Select
        ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
        SUV_Range = Sheet_Name & "!$B$2:$B$" & row_total & "," & Sheet_Name & "!$" & Col_lett(t) & "$2:$" & Col_lett(t) & "$" & row_total
        With ActiveChart
            .SetSourceData Source:=Range(SUV_Range)
            .ChartTitle.Text = Range(Col_lett(t) & "1").Value
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time (mins)"
            .Axes(xlCategory).MinimumScale = 0
            .Axes(xlCategory).MaximumScale = 150
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "SUVr"
        End With
        t = t + 5
    Loop
    
    u = 8
    Do While u < column_total_after_SUVr_and_Time
        Range("B3:B" & row_total).Select
        Range(Col_lett(u) & "3:" & Col_lett(u) & row_total).Select
        ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
        SUV_Range = Sheet_Name & "!$B$3:$B$" & row_total & "," & Sheet_Name & "!$" & Col_lett(u) & "$3:$" & Col_lett(u) & "$" & row_total
        With ActiveChart
            .SetSourceData Source:=Range(SUV_Range)
            .ChartTitle.Text = Range(Col_lett(u) & "1").Value
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time (mins)"
            .Axes(xlCategory).MinimumScale = 0
            .Axes(xlCategory).MaximumScale = 150
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "SUVr"
        End With
        u = u + 5
    Loop

    ' Autofit columns 1 to column_total_after_SUV
    Dim Z As Long
    For Z = 1 To column_total_after_SUV
        Columns(Col_lett(Z)).EntireColumn.AutoFit
    Next Z

    ' Remove formulas from the top row (convert to values)
    Rows("1:1").Copy
    Rows("1:1").PasteSpecial Paste:=xlPasteValues

    '==============================
    ' Chart Lineup
    '==============================
    Chart_count = ActiveSheet.ChartObjects.Count
    Width = 200
    Height = 150
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
        ActiveSheet.ChartObjects(Chart_index).Top = 770
    Next Chart_index
    For Chart_index = ((2 * Chart_count) / 3) + 1 To Chart_count
        ActiveSheet.ChartObjects(Chart_index).Top = 920
    Next Chart_index

    '==============================
    ' SUV and SUVr Multiplots
    '==============================
    ROIs = Array( _
        Array("SUV_W_MUBADA_202_Wbrain", "SUV_Hippocampus_L_AAL", "SUV_Hippocampus_R_AAL", "SUV_Precuneus_L_AAL", "SUV_Precuneus_R_AAL", "SUV_Putamen_L_AAL", "SUV_Putamen_R_AAL", "SUV_blcere_all"), _
        Array("SUV_W_MUBADA_202_Wbrain", "SUV_Hippocampus_L_AAL", "SUV_Hippocampus_R_AAL", "SUV_Precuneus_L_AAL", "SUV_Precuneus_R_AAL", "SUV_Putamen_L_AAL", "SUV_Putamen_R_AAL", "SUV_cerebellum_crus1_v5"), _
        Array("SUVr(blcere)_W_MUBADA_202_Wbrain", "SUVr(blcere)_Hippocampus_L_AAL", "SUVr(blcere)_Hippocampus_R_AAL", "SUVr(blcere)_Precuneus_L_AAL", "SUVr(blcere)_Precuneus_R_AAL", "SUVr(blcere)_Putamen_L_AAL", "SUVr(blcere)_Putamen_R_AAL", "SUVr(blcere)_blcere_all"), _
        Array("SUVr(Cerecrus)_W_MUBADA_202_Wbrain", "SUVr(Cerecrus)_Hippocampus_L_AAL", "SUVr(Cerecrus)_Hippocampus_R_AAL", "SUVr(Cerecrus)_Precuneus_L_AAL", "SUVr(Cerecrus)_Precuneus_R_AAL", "SUVr(Cerecrus)_Putamen_L_AAL", "SUVr(Cerecrus)_Putamen_R_AAL", "SUVr(Cerecrus)_cerebellum_crus1_v5"))
    ROIs_length = UBound(ROIs)
    
    For r = 0 To ROIs_length
        Sub_ROIs_length = UBound(ROIs(r))
        For rr = 0 To Sub_ROIs_length
            Range("C1").Select
            Cells.Find(What:=ROIs(r)(rr), LookIn:=xlFormulas, LookAt:=xlPart, _
                SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=True).Activate
            Col_array(r, rr) = Col_lett(Selection.Column)
        Next rr
        
        With Range("B2:B" & row_total)
            Multi_range = Col_array(r, 0) & "2:" & Col_array(r, 0) & row_total & "," & _
                          Col_array(r, 1) & "2:" & Col_array(r, 1) & row_total & "," & _
                          Col_array(r, 2) & "2:" & Col_array(r, 2) & row_total & "," & _
                          Col_array(r, 3) & "2:" & Col_array(r, 3) & row_total & "," & _
                          Col_array(r, 4) & "2:" & Col_array(r, 4) & row_total & "," & _
                          Col_array(r, 5) & "2:" & Col_array(r, 5) & row_total & "," & _
                          Col_array(r, 6) & "2:" & Col_array(r, 6) & row_total & "," & _
                          Col_array(r, 7) & "2:" & Col_array(r, 7) & row_total & "," & _
                          Col_array(r, 8) & "2:" & Col_array(r, 8) & row_total
        End With
        
        ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmooth).Select
        Multiplot = Sheet_Name & "!$B$2:$B$" & row_total & "," & _
                     Sheet_Name & "!$" & Col_array(r, 0) & "$2:$" & Col_array(r, 0) & "$" & row_total & "," & _
                     Sheet_Name & "!$" & Col_array(r, 1) & "$2:$" & Col_array(r, 1) & "$" & row_total & "," & _
                     Sheet_Name & "!$" & Col_array(r, 2) & "$2:$" & Col_array(r, 2) & "$" & row_total & "," & _
                     Sheet_Name & "!$" & Col_array(r, 3) & "$2:$" & Col_array(r, 3) & "$" & row_total & "," & _
                     Sheet_Name & "!$" & Col_array(r, 4) & "$2:$" & Col_array(r, 4) & "$" & row_total & "," & _
                     Sheet_Name & "!$" & Col_array(r, 5) & "$2:$" & Col_array(r, 5) & "$" & row_total & "," & _
                     Sheet_Name & "!$" & Col_array(r, 6) & "$2:$" & Col_array(r, 6) & "$" & row_total & "," & _
                     Sheet_Name & "!$" & Col_array(r, 7) & "$2:$" & Col_array(r, 7) & "$" & row_total & "," & _
                     Sheet_Name & "!$" & Col_array(r, 8) & "$2:$" & Col_array(r, 8) & "$" & row_total
        With ActiveChart
            .SetSourceData Source:=Range(Multiplot)
            .ChartType = xlXYScatterSmooth
            .HasLegend = True
            With .SeriesCollection(1)
                .Name = "MUBADA"
                .Format.Fill.ForeColor.RGB = RGB(19, 149, 186)
                .Format.Line.ForeColor.RGB = RGB(19, 149, 186)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            With .SeriesCollection(2)
                .Name = "Hippocampus_L"
                .Format.Fill.ForeColor.RGB = RGB(13, 60, 85)
                .Format.Line.ForeColor.RGB = RGB(13, 60, 85)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            With .SeriesCollection(3)
                .Name = "Hippocampus_R"
                .Format.Fill.ForeColor.RGB = RGB(192, 46, 29)
                .Format.Line.ForeColor.RGB = RGB(192, 46, 29)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            With .SeriesCollection(4)
                .Name = "Precuneus_L"
                .Format.Fill.ForeColor.RGB = RGB(241, 108, 32)
                .Format.Line.ForeColor.RGB = RGB(241, 108, 32)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            With .SeriesCollection(5)
                .Name = "Precuneus_R"
                .Format.Fill.ForeColor.RGB = RGB(239, 139, 44)
                .Format.Line.ForeColor.RGB = RGB(239, 139, 44)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            With .SeriesCollection(6)
                .Name = "Putamen_L"
                .Format.Fill.ForeColor.RGB = RGB(235, 200, 68)
                .Format.Line.ForeColor.RGB = RGB(235, 200, 68)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            With .SeriesCollection(7)
                .Name = "Putamen_R"
                .Format.Fill.ForeColor.RGB = RGB(162, 184, 108)
                .Format.Line.ForeColor.RGB = RGB(162, 184, 108)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            With .SeriesCollection(8)
                .Name = "Cerebellum"
                .Format.Fill.ForeColor.RGB = RGB(92, 167, 147)
                .Format.Line.ForeColor.RGB = RGB(92, 167, 147)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            .ChartTitle.Text = ROIs(r)(8)
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time (mins)"
            .Axes(xlCategory).MinimumScale = 0
            .Axes(xlCategory).MaximumScale = 150
            .Axes(xlValue, xlPrimary).HasTitle = True
            If r < 2 Then
                .Axes(xlValue, xlPrimary).AxisTitle.Text = "SUV"
                .Axes(xlValue).MinimumScale = SUV_min
                .Axes(xlValue).MaximumScale = SUV_max
            Else
                .Axes(xlValue, xlPrimary).AxisTitle.Text = "SUVr"
                .Axes(xlValue).MinimumScale = SUVr_min
                .Axes(xlValue).MaximumScale = SUVr_max
            End If
        End With
        
        Multiplot_index = ActiveSheet.ChartObjects.Count
        ActiveSheet.ChartObjects(Multiplot_index).Width = 600
        ActiveSheet.ChartObjects(Multiplot_index).Height = 350
        Select Case r
            Case 0
                ActiveSheet.ChartObjects(Multiplot_index).Top = 2500
                ActiveSheet.ChartObjects(Multiplot_index).Left = 0
            Case 1
                ActiveSheet.ChartObjects(Multiplot_index).Top = 2500
                ActiveSheet.ChartObjects(Multiplot_index).Left = 600
            Case 2
                ActiveSheet.ChartObjects(Multiplot_index).Top = 2850
                ActiveSheet.ChartObjects(Multiplot_index).Left = 0
            Case 3
                ActiveSheet.ChartObjects(Multiplot_index).Top = 2850
                ActiveSheet.ChartObjects(Multiplot_index).Left = 600
        End Select
        ActiveChart.ChartStyle = 241
    Next r

    Range("A70").Select
    ActiveWindow.Zoom = 80
    ActiveSheet.Name = Subject_study

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Neuro Macro Error"
    Resume Cleanup
End Sub
