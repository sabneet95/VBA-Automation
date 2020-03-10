Public aborting_mechanism As Integer
Function onlyDigits(s As String) As String
    Dim retval As String
    Dim i As Integer
    retval = ""
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next
    onlyDigits = retval
End Function
Function Col_lett(ByVal ColumnNumber As Integer)
Col_lett = Replace(Replace(Cells(1, ColumnNumber).Address, "1", ""), "$", "")
End Function

Sub test()

'''''''''''''''''''''''''''''''''''Variable Declarations''''''''''''''''''''''''''''''
    Dim Initial_frames As String
    Dim Subject_frame_number As String
    Dim Subject_study As String
    Dim Initial_delay_divide As Integer
    Dim ROIs As Variant
    Dim Address(10) As Variant
    'Dim Address2(10) As Variant
    'Dim Average_label As String
    Dim Weight As Double
    Dim Dose As Double
    Dim SUV_formula As String
    Dim Decay_SUV_formula As String
    Dim SUVr_Formula As String
    Dim MAX_formula As String
    Dim MAX_MAX_formula As String
    Dim MIN_formula As String
    Dim MIN_MIN_formula As String
    Dim Sheet_Name As String
    Dim Width As Single, Height As Single
    Dim NumWide As Long
    Dim Chart_index As Long, Chart_count As Long
    Dim Col_array(3, 8) As Variant
    Dim Multi_range As String
    Dim vAxis As Variant
    Dim Multiplot As String
    
''''''''''''''''''''''''''''''''''''''File Upgrade''''''''''''''''''''''''''''''''''''

    Current_path = Application.ActiveWorkbook.FullName
    Extension = Right(Current_path, 1)
    If Not Extension = "x" Then
        Upgraded_File = Current_path + "x"
        ActiveWorkbook.SaveAs Filename:=Upgraded_File, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        Workbooks.Open Filename:=Upgraded_File
    End If
    Range("A1").Select
    If InStr(Range("A2").Value, "mean") > 0 Then '''''''''''Delete Mean Frame'''''''''
        Rows("2:2").Select
        Selection.Delete Shift:=xlUp
    End If

''''''''''''''''''''''''''''''''''''Frame Sequence Fix''''''''''''''''''''''''''''''''
''''''''''''''''''''This is a very crude and hard coded way to do it''''''''''''''''''

    Frames = MsgBox("Would you like to rearrange the frames ascendingly by time?", vbYesNo + vbQuestion, "Frame Order")
    If Frames = vbYes Then
'        Rows("3:5").Select
'        Selection.Cut
'        Range("A45").Select
'        ActiveSheet.Paste
'        Rows("6:13").Select
'        Selection.Cut
'        Range("A3").Select
'        ActiveSheet.Paste
'        Rows("45:47").Select
'        Selection.Cut
'        Range("A11").Select
'        ActiveSheet.Paste
'        Rows("2:13").Select
'        Selection.Cut
'        Range("A45").Select
'        ActiveSheet.Paste
'        Rows("15:15").Select
'        Rows("15:24").Select
'        Selection.Cut
'        Range("A60").Select
'        ActiveSheet.Paste
'        Rows("26:29").Select
'        Selection.Cut
'        Range("A70").Select
'        ActiveSheet.Paste
'        Rows("25:25").Select
'        Rows("14:14").Select
'        Selection.Cut
'        Range("A2").Select
'        ActiveSheet.Paste
'        Rows("25:25").Select
'        Selection.Cut
'        Range("A3").Select
'        ActiveSheet.Paste
'        Rows("30:36").Select
'        Selection.Cut
'        Range("A4").Select
'        ActiveSheet.Paste
'        Rows("60:73").Select
'        Selection.Cut
'        Range("A11").Select
'        ActiveSheet.Paste
'        Rows("45:56").Select
'        Selection.Cut
'        Range("A25").Select
'        ActiveSheet.Paste

        column_total = Range("A1").End(xlToRight).Offset(0, 4).Column
        row_total = Range("A1").End(xlDown).Offset(0, 1).Row
        
        With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
            .SortFields.Add Key:=Range("A2"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            .SetRange Range("A2:" & Col_lett(column_total) & row_total)
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    
        With Range("A1:A" & row_total)
            .Replace What:=".img", Replacement:="", LookAt:=xlPart, SearchOrder _
                :=xlByColumns, MatchCase:=False, SearchFormat:=False, ReplaceFormat:= _
                False
            .Replace What:="wrrxx", Replacement:="", LookAt:=xlPart, SearchOrder _
                :=xlByColumns, MatchCase:=False, SearchFormat:=False, ReplaceFormat:= _
                False
            .Find(What:="*_d*_f", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
                :=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:= _
                False, SearchFormat:=False).Activate
            Subject_value = ActiveCell.Value
            Subject_length = Len(Subject_value)
            Subject_underscore = InStr(UCase(Subject_value), "_D")
            Subject_frame_number = Right(Subject_value, Subject_length - Subject_underscore)
            Subject_study = Left(Subject_value, (Subject_length - Len(Subject_frame_number)) - 1)
            Intial_frames = Subject_study & "_f"
            
        Set C = .Find("*_d*_f", LookIn:=xlValues)
            Do While Not C Is Nothing
                Subject_value = C.Value
                Subject_length = Len(Subject_value)
                Subject_underscore = InStr(UCase(Subject_value), "_D")
                Subject_frame_number = Right(Subject_value, Subject_length - Subject_underscore)
                C.Value = onlyDigits(Subject_frame_number)
                Set C = .FindNext(C)
            Loop
    
        Set D = .Find(Intial_frames, LookIn:=xlValues)
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
        
        For D = 2 To Initial_delay_divide - 1
            Digit_value = Range("A" & D).Value
            Range("A" & D).Value = Subject_study & "_Frame" & Digit_value
        Next
        
        For C = Initial_delay_divide To row_total
            Digit_value = Range("A" & C).Value
            Range("A" & C).Value = Subject_study & "_Delayed_Frame" & Digit_value
        Next
    End If
    
'''''''''''''''''''''''''''''''''''''''ROIs Creation''''''''''''''''''''''''''''''''''
    
    ROIs = Array(Array("Lingual_L_AAL.nii", "Occipital_Sup_L_AAL.nii", "Occipital_Mid_L_AAL.nii", "Occipital_Inf_L_AAL.nii", "Cuneus_L_AAL.nii", "Calcarine_L_AAL.nii") _
         , Array("Lingual_R_AAL.nii", "Occipital_Sup_R_AAL.nii", "Occipital_Mid_R_AAL.nii", "Occipital_Inf_R_AAL.nii", "Cuneus_R_AAL.nii", "Calcarine_R_AAL.nii") _
         , Array("Angular_L_AAL.nii", "SupraMarginal_L_AAL.nii", "Parietal_Sup_L_AAL.nii", "Parietal_Inf_L_AAL.nii", "Precuneus_L_AAL.nii") _
         , Array("Angular_R_AAL.nii", "SupraMarginal_R_AAL.nii", "Parietal_Sup_R_AAL.nii", "Parietal_Inf_R_AAL.nii", "Precuneus_R_AAL.nii") _
         , Array("Temporal_Pole_Mid_L_AAL.nii", "Temporal_Sup_L_AAL.nii", "Temporal_Pole_Mid_L_AAL.nii", "Temporal_Mid_L_AAL.nii", "Temporal_Inf_L_AAL.nii") _
         , Array("Temporal_Pole_Mid_R_AAL.nii", "Temporal_Sup_R_AAL.nii", "Temporal_Pole_Mid_R_AAL.nii", "Temporal_Mid_R_AAL.nii", "Temporal_Inf_R_AAL.nii") _
         , Array("Frontal_Sup_L_AAL.nii", "Frontal_Mid_L_AAL.nii", "Frontal_Inf_Oper_L_AAL.nii", "Frontal_Inf_Tri_L_AAL.nii", "Frontal_Sup_Medial_L_AAL.nii", "Supp_Motor_Area_L_AAL.nii") _
         , Array("Frontal_Sup_R_AAL.nii", "Frontal_Mid_R_AAL.nii", "Frontal_Inf_Oper_R_AAL.nii", "Frontal_Inf_Tri_R_AAL.nii", "Frontal_Sup_Medial_R_AAL.nii", "Supp_Motor_Area_R_AAL.nii"))
    
    ROIs_length = UBound(ROIs)
    ROIs_names = Array("Occipital_L", "Occipital_R", "Parietal_L", "Parietal_R", "Temporal_L", "Temporal_R", "Frontal_L", "Frontal_R")

    For r = 0 To ROIs_length
        Sub_ROIs_length = UBound(ROIs(r))
        
        If r > 0 Then
            Range("A1").End(xlToRight).Offset(0, 1).Value = "Keep Blank!"
        End If
        
        Range("A1").End(xlToRight).Offset(0, 1).Select
        column_total = ActiveCell.Column
        column_addition = ActiveCell.Address
        Range("A1").End(xlDown).Offset(0, 1).Select
        row_total = ActiveCell.Row
    
        For rr = 0 To Sub_ROIs_length
            Range("A1").Select
            Cells.Find(What:=ROIs(r)(rr), After:=ActiveCell, LookIn:= _
                xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                xlNext, MatchCase:=False, SearchFormat:=False).Offset(1, 0).Activate
            'Average_label = Average_label + " & Address(" & rr & ") & " & Chr(34) & "," & Chr(34)
            Address(rr) = ActiveCell.Address(False, False)
            'Address2(rr) = ActiveCell.Offset(0, 1).Address(False, False)
        Next rr
    
        Range(column_addition).Value = ROIs_names(r)
        Range(column_addition).Offset(1, 0).Select
        
        If Sub_ROIs_length = 4 Then
            Selection.Value = "=AVERAGE(" & Address(0) & "," & Address(1) & "," & Address(2) & "," & Address(3) & "," & Address(4) & ")"
            Selection.AutoFill Destination:=Range(ActiveCell.Address(False, False), Col_lett(ActiveCell.Column) & row_total), Type:=xlFillDefault
        ElseIf Sub_ROIs_length = 5 Then
            Selection.Value = "=AVERAGE(" & Address(0) & "," & Address(1) & "," & Address(2) & "," & Address(3) & "," & Address(4) & "," & Address(5) & ")"
            Selection.AutoFill Destination:=Range(ActiveCell.Address(False, False), Col_lett(ActiveCell.Column) & row_total), Type:=xlFillDefault
        End If

'                                              Weighted Average Method
'        If Sub_ROIs_length = 4 Then
'            Selection.Value = "=(" & Address(0) & "*" & Address2(0) & "+" & Address(1) & "*" & Address2(1) & "+" & Address(2) & "*" & Address2(2) & "+" & Address(3) & "*" & Address2(3) & "+" & Address(4) & "*" & Address2(4) & ")/(" & Address2(0) & "+" & Address2(1) & "+" & Address2(2) & "+" & Address2(3) & "+" & Address2(4) & ")"
'            Selection.AutoFill Destination:=Range(ActiveCell.Address(False, False), Col_lett(ActiveCell.Column) & row_total), Type:=xlFillDefault
'        ElseIf Sub_ROIs_length = 5 Then
'            'Selection.Value = "=(" & Address(0) & "*" & Address2(0) & "+" & Address(1) & "*" & Address2(1) & "+" & Address(2) & "*" & Address2(2) & "+" & Address(3) & "*" & Address2(3) & "+" & Address(4) & "*" & Address2(4) & "+" & Address(5) & "*" & Address2(5) & ")/(" & Address2(0) & "+" & Address2(1) & "+" & Address2(2) & "+" & Address2(3) & "+" & Address2(4) "+" & Address2(5) & ")"
'            Selection.AutoFill Destination:=Range(ActiveCell.Address(False, False), Col_lett(ActiveCell.Column) & row_total), Type:=xlFillDefault
'        End If
        
        If r = ROIs_length Then
            Range("A1").End(xlToRight).Offset(0, 1).Value = "Keep Blank!"
        End If
    Next r
    
'''''''''''''''''''''''''''''''Rows And Columns count'''''''''''''''''''''''''''''''''

    Range("A1").End(xlToRight).Offset(0, 4).Select
    column_total = ActiveCell.Column
    Range("A1").End(xlDown).Offset(0, 1).Select
    row_total = ActiveCell.Row

''''''''''''''''''''''''''''''''''''''Formatting''''''''''''''''''''''''''''''''''''''

    Range("A1").Select
    Cells.Replace What:=".img", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByColumns, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Cells.Replace What:="wrrxx", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByColumns, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Columns(1).EntireColumn.AutoFit
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .Color = 6684876
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True

    Range(Col_lett(2) & "1" & ":" & Col_lett(column_total - 3) & "1").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13421619
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True

'''''''''''''''''''''''''''''''''''Time Columns Creation''''''''''''''''''''''''''''''
'                                        Method 1:
'
'    Times = MsgBox("Would you like to continue with the default times?", vbYesNo + vbQuestion, "SUV and SUVr graphs")
'    If Times = vbYes Then
'
'        Columns(1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'        Range(Col_lett(1) + "1").Value = "Time Intervals"
'        Columns(2).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'        Range(Col_lett(2) + "1").Value = "Start Time"
'
'        Range("A2").Value = "0.5"
'        Range("A2").Select
'        Selection.AutoFill Destination:=Range("A2:A7"), Type:=xlFillDefault
'        Range("A8").Value = "1"
'        Range("A8").Select
'        Selection.AutoFill Destination:=Range("A8:A11"), Type:=xlFillDefault
'        Range("A12").Value = "2"
'        Range("A12").Select
'        Selection.AutoFill Destination:=Range("A12:A15"), Type:=xlFillDefault
'        Range("A16").Value = "5"
'        Range("A16").Select
'        Selection.AutoFill Destination:=Range("A16:A36"), Type:=xlFillDefault
'
'        Range("B2").Value = "0"
'        Range("B3").Value = "=R[-1]C+0.5"
'        Range("B3").Select
'        Selection.AutoFill Destination:=Range("B3:B8"), Type:=xlFillDefault
'        Range("B9").Value = "=R[-1]C+1"
'        Range("B9").Select
'        Selection.AutoFill Destination:=Range("B9:B12"), Type:=xlFillDefault
'        Range("B13").Value = "=R[-1]C+2"
'        Range("B13").Select
'        Selection.AutoFill Destination:=Range("B13:B16"), Type:=xlFillDefault
'        Range("B17").Value = "=R[-1]C+5"
'        Range("B17").Select
'        Selection.AutoFill Destination:=Range("B17:B24"), Type:=xlFillDefault
'        Range("B25").Value = "90"
'        Range("B26").Value = "=R[-1]C+5"
'        Range("B26").Select
'        Selection.AutoFill Destination:=Range("B26:B36"), Type:=xlFillDefault
'    Else
'        Exit Sub
'    End If
'
'                                   Method 2:
    
    Columns(1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range(Col_lett(1) + "1").Value = "Time Intervals"
    Columns(2).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range(Col_lett(2) + "1").Value = "Start Time"
    
    UserForm1.Show
    If Flag > 0 Then
        Exit Sub
    End If
    
    Range("B1").Select
    Cells.Find(What:="90", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    Decay_row = ActiveCell.Row

''''''''''''''''''''''''''''''''''Weight And Dose'''''''''''''''''''''''''''''''''''''

    Weight = InputBox("Now, please proceed to MIM and acquire the subject weight in Kg (by design Kg -> g):", ["Subject Weight"])
    Weight_in_grams = Weight * 1000#
    Dose = InputBox("And now, input the total dose registered in mCi (by design mCi -> Bq):")
    Dose_in_Bq = Dose * 37000000#
    Range("A1").End(xlDown).Offset(2, 0).Value = "Patient Weight:"
    Range("B1").End(xlDown).Offset(2, 0).Value = "Total Dose:"
    Range("A1").End(xlDown).Offset(3, 0).Select
    Selection.Value = Weight_in_grams & " g"
    Selection.ClearComments
    Selection.AddComment
    Selection.Comment.Text Text:="Please, keep the same format when modifying as it might accidentally break the rest of the functions! For your reference, copy the following default, if in vain: 10000 g"
    Selection.Comment.Visible = True
    Selection.Comment.Shape.Select True
    Selection.ShapeRange.ScaleWidth 1.5, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.9, msoFalse, msoScaleFromTopLeft
    Range("A1").End(xlDown).Offset(3, 0).Select
    Selection.Comment.Visible = False
    Weight_row = Selection.Row

    Range("B1").End(xlDown).Offset(3, 0).Select
    Selection.Value = Dose_in_Bq & " Bq"
    Selection.ClearComments
    Selection.AddComment
    Selection.Comment.Text Text:="Same here, the default is: 11100000 Bq"
    Selection.Comment.Visible = True
    Selection.Comment.Shape.Select True
    Selection.ShapeRange.ScaleWidth 1, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.9, msoFalse, msoScaleFromTopLeft
    Range("B1").End(xlDown).Offset(3, 0).Select
    Selection.Comment.Visible = False
    Dose_row = Selection.Row

'''''''''''''''''''''''''''''SUV Column Creation And Calculations'''''''''''''''''''''

    i = 6
    Do While i < column_total
        Columns(i).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range(Col_lett(i) + "1").Value = "=CONCATENATE(""SUV_"",RC[-2])"

        With Range(Col_lett(i) + "1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 49407
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Range(Col_lett(i) + "1").Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        Range(Col_lett(i) + "1").Font.Bold = True
        SUV_formula = "=RC[-2]*(LEFT(R" & Weight_row & "C1,LEN(R" & Weight_row & "C1)-2))/(LEFT(R" & Dose_row & "C2,LEN(R" & Dose_row & "C2)-3))"
        Decay_SUV_formula = "=RC[-2]*(LEFT(R" & Weight_row & "C1,LEN(R" & Weight_row & "C1)-2))/((LEFT(R" & Dose_row & "C2,LEN(R" & Dose_row & "C2)-3))*EXP(-0.693*90/109.77))"
        Range(Col_lett(i) + "2").FormulaR1C1 = SUV_formula
        Range(Col_lett(i) & Decay_row).FormulaR1C1 = Decay_SUV_formula
        Range(Col_lett(i) + "2").AutoFill Destination:=Range(Col_lett(i) & "2" & ":" & Col_lett(i) & Decay_row - 1), Type:=xlFillDefault
        Range(Col_lett(i) & Decay_row).AutoFill Destination:=Range(Col_lett(i) & Decay_row & ":" & Col_lett(i) & row_total), Type:=xlFillDefault

        i = i + 3
        column_total = column_total + 1
    Loop

'''''''''''''''''''''''''''''SUVr Column Creation And Calculations'''''''''''''''''''''

    Range("A1").End(xlToRight).Offset(0, 3).Select
    column_total_after_SUV = ActiveCell.Column
    Range(Col_lett(1) & "1").Select 'Lets change the view
    j = 7
    Do While j < column_total_after_SUV
        Columns(j).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range(Col_lett(j) + "1").Value = "=CONCATENATE(""SUVr(blcere)_"",RC[-3])"

        With Range(Col_lett(j) + "1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 49407
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Range(Col_lett(j) + "1").Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        Range(Col_lett(j) + "1").Font.Bold = True

        j = j + 4
        column_total_after_SUV = column_total_after_SUV + 1
    Loop

    j = 8
    Do While j < column_total_after_SUV
        Columns(j).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range(Col_lett(j) + "1").Value = "=CONCATENATE(""SUVr(Cerecrus)_"",RC[-4])"

        With Range(Col_lett(j) + "1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 49407
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Range(Col_lett(j) + "1").Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        Range(Col_lett(j) + "1").Font.Bold = True

        j = j + 5
        column_total_after_SUV = column_total_after_SUV + 1
    Loop

    Range("A1").Select

    Cells.Find(What:="blcere_all.nii", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Offset(2, 2).Activate
    Cerebellum = Selection.Column

    Cells.Find(What:="cerebellum_crus1_v5.nii", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Offset(2, 2).Activate
    Cerebellum_Crus = Selection.Column

    k = 7
    Do While k < column_total_after_SUV
        SUVr_Formula = "=RC[-1]/R[0]C" & Cerebellum
        Range(Col_lett(k) + "2").FormulaR1C1 = SUVr_Formula
        Range(Col_lett(k) + "2").AutoFill Destination:=Range(Col_lett(k) & "2" & ":" & Col_lett(k) & row_total), Type:=xlFillDefault

        k = k + 5
    Loop

    k = 8
    Do While k < column_total_after_SUV
        SUVr_Formula = "=RC[-2]/R[0]C" & Cerebellum_Crus
        Range(Col_lett(k) + "2").FormulaR1C1 = SUVr_Formula
        Range(Col_lett(k) + "2").AutoFill Destination:=Range(Col_lett(k) & "2" & ":" & Col_lett(k) & row_total), Type:=xlFillDefault

        k = k + 5
    Loop

    Cells.Replace What:=".nii", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

''''''''''''''''''''''''''''''''''''''Graphs Creation''''''''''''''''''''''''''''''''''''

        Range("A1").End(xlToRight).Offset(0, 2).Select
        column_total_after_SUVr_and_Time = ActiveCell.Column
        Range("A1").End(xlDown).Offset(0, 1).Select
        row_total = ActiveCell.Row

        P = 6
        Do While P < column_total_after_SUVr_and_Time
            MAX_formula = "=MAX(R[-" & row_total + 3 & "]C:R[-1]C)"
            MIN_formula = "=MIN(R[-" & row_total + 4 & "]C:R[-2]C)"
            Range(Col_lett(P) & row_total + 5).Value = MAX_formula
            Range(Col_lett(P) & row_total + 6).Value = MIN_formula
            P = P + 5
        Loop

        q = 7
        Do While q < column_total_after_SUVr_and_Time
            MAX_formula = "=MAX(R[-" & row_total + 5 & "]C:R[-1]C)"
            MIN_formula = "=MIN(R[-" & row_total + 6 & "]C:R[-2]C)"
            Range(Col_lett(q) & row_total + 7).Value = MAX_formula
            Range(Col_lett(q) & row_total + 8).Value = MIN_formula
            q = q + 5
        Loop

        r = 8
        Do While r < column_total_after_SUVr_and_Time
            MAX_formula = "=MAX(R[-" & row_total + 7 & "]C:R[-1]C)"
            MIN_formula = "=MIN(R[-" & row_total + 8 & "]C:R[-2]C)"
            Range(Col_lett(r) & row_total + 9).Value = MAX_formula
            Range(Col_lett(r) & row_total + 10).Value = MIN_formula
            r = r + 5
        Loop

        MAX_MAX_formula = "=MAX(RC[-" & column_total - 8 & "]:RC[-1])"
        MIN_MIN_formula = "=MIN(RC[-" & column_total - 8 & "]:RC[-1])"

        Range(Col_lett(column_total_after_SUVr_and_Time) & row_total).Offset(5, -2).Value = MAX_MAX_formula
        SUV_max = WorksheetFunction.RoundUp(Range(Col_lett(column_total_after_SUVr_and_Time) & row_total).Offset(5, -2).Value, 1)

        Range(Col_lett(column_total_after_SUVr_and_Time) & row_total).Offset(6, -2).Value = MIN_MIN_formula
        SUV_min = WorksheetFunction.RoundDown(Range(Col_lett(column_total_after_SUVr_and_Time) & row_total).Offset(6, -2).Value, 1)

        Range(Col_lett(column_total_after_SUVr_and_Time) & row_total).Offset(7, -2).Value = MAX_MAX_formula
        SUVr_max = WorksheetFunction.RoundUp(Range(Col_lett(column_total_after_SUVr_and_Time) & row_total).Offset(7, -2).Value, 1)

        Range(Col_lett(column_total_after_SUVr_and_Time) & row_total).Offset(8, -2).Value = MIN_MIN_formula
        SUVr_min = WorksheetFunction.RoundDown(Range(Col_lett(column_total_after_SUVr_and_Time) & row_total).Offset(8, -2).Value, 1)

        Range(Col_lett(column_total_after_SUVr_and_Time) & row_total).Offset(9, -2).Value = MAX_MAX_formula
        SUVr2_max = WorksheetFunction.RoundUp(Range(Col_lett(column_total_after_SUVr_and_Time) & row_total).Offset(9, -2).Value, 1)

        Range(Col_lett(column_total_after_SUVr_and_Time) & row_total).Offset(10, -2).Value = MIN_MIN_formula
        SUVr2_min = WorksheetFunction.RoundDown(Range(Col_lett(column_total_after_SUVr_and_Time) & row_total).Offset(10, -2).Value, 1)

        SUVr_max = WorksheetFunction.Max(SUVr_max, SUVr2_max)
        SUVr_min = WorksheetFunction.Min(SUVr_min, SUVr2_min)


        Rows(row_total + 5 & ":" & row_total + 10).Select
        Selection.Delete Shift:=xlUp
        Sheet_Name = ActiveSheet.Name

        s = 6
        Do While s < column_total_after_SUVr_and_Time
            Range("B2:B" & row_total).Select
            Range(Col_lett(s) & "2:" & Col_lett(s) & row_total).Select
            ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
            SUV_Range = Sheet_Name & "!$B$2:$B$" & row_total & "," & Sheet_Name & "!$" & Col_lett(s) & "$2:$" & Col_lett(s) & "$" & row_total
            With ActiveChart
                .SetSourceData Source:=Range(SUV_Range)
                .ChartTitle.Text = Range(Col_lett(s) & 1).Value
                
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (mins)"
                .Axes(xlCategory).Format.Line.EndArrowheadStyle = msoArrowheadStealth
                .Axes(xlCategory).Format.Line.EndArrowheadLength = msoArrowheadLong
                .Axes(xlCategory).Format.Line.EndArrowheadWidth = msoArrowheadWide
                .Axes(xlCategory).MinimumScale = 0
                .Axes(xlCategory).MaximumScale = 150
                .Axes(xlCategory).MajorUnit = 30
                .Axes(xlCategory).MajorTickMark = xlCross
                .Axes(xlCategory).MinorTickMark = xlCross
                .Axes(xlCategory).MajorTickMark = xlCross
                
                .Axes(xlValue, xlPrimary).HasTitle = True
                .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "SUV"
                .Axes(xlValue).Format.Line.EndArrowheadStyle = msoArrowheadStealth
                .Axes(xlValue).Format.Line.EndArrowheadLength = msoArrowheadLong
                .Axes(xlValue).Format.Line.EndArrowheadWidth = msoArrowheadWide
                .Axes(xlValue).MinimumScale = SUV_min
                .Axes(xlValue).MaximumScale = SUV_max
                .Axes(xlValue).MajorTickMark = xlCross

                .SetElement (msoElementPrimaryValueGridLinesNone)
                .SetElement (msoElementPrimaryCategoryGridLinesNone)
                
                .FullSeriesCollection(1).MarkerForegroundColorIndex = -4142
                .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 102, 0)
                .FullSeriesCollection(1).MarkerSize = 4
                .FullSeriesCollection(1).MarkerStyle = 8
                
                .ChartArea.Font.Color = RGB(31, 78, 121)
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
            s = s + 5
        Loop
        t = 7
        Do While t < column_total_after_SUVr_and_Time
            Range("B2:B" & row_total).Select
            Range(Col_lett(t) & "2:" & Col_lett(t) & row_total).Select
            ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
            
            SUV_Range = Sheet_Name & "!$B$2:$B$" & row_total & "," & Sheet_Name & "!$" & Col_lett(t) & "$2:$" & Col_lett(t) & "$" & row_total
            With ActiveChart
                .SetSourceData Source:=Range(SUV_Range)
                .ChartTitle.Text = Range(Col_lett(t) & 1).Value
                
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (mins)"
                .Axes(xlCategory).Format.Line.EndArrowheadStyle = msoArrowheadStealth
                .Axes(xlCategory).Format.Line.EndArrowheadLength = msoArrowheadLong
                .Axes(xlCategory).Format.Line.EndArrowheadWidth = msoArrowheadWide
                .Axes(xlCategory).MinimumScale = 0
                .Axes(xlCategory).MaximumScale = 150
                .Axes(xlCategory).MajorUnit = 30
                .Axes(xlCategory).MajorTickMark = xlCross
                .Axes(xlCategory).MinorTickMark = xlCross
                .Axes(xlCategory).MajorTickMark = xlCross
                
                .Axes(xlValue, xlPrimary).HasTitle = True
                .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "SUVr"
                .Axes(xlValue).Format.Line.EndArrowheadStyle = msoArrowheadStealth
                .Axes(xlValue).Format.Line.EndArrowheadLength = msoArrowheadLong
                .Axes(xlValue).Format.Line.EndArrowheadWidth = msoArrowheadWide
                .Axes(xlValue).MinimumScale = SUVr_min
                .Axes(xlValue).MaximumScale = SUVr_max
                .Axes(xlValue).MajorTickMark = xlCross

                .SetElement (msoElementPrimaryValueGridLinesNone)
                .SetElement (msoElementPrimaryCategoryGridLinesNone)
                
                .FullSeriesCollection(1).MarkerForegroundColorIndex = -4142
                .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 204, 0)
                .FullSeriesCollection(1).MarkerSize = 4
                .FullSeriesCollection(1).MarkerStyle = 8
                
                .ChartArea.Font.Color = RGB(31, 78, 121)
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
                .ChartTitle.Text = Range(Col_lett(u) & 1).Value
                
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (mins)"
                .Axes(xlCategory).Format.Line.EndArrowheadStyle = msoArrowheadStealth
                .Axes(xlCategory).Format.Line.EndArrowheadLength = msoArrowheadLong
                .Axes(xlCategory).Format.Line.EndArrowheadWidth = msoArrowheadWide
                .Axes(xlCategory).MinimumScale = 0
                .Axes(xlCategory).MaximumScale = 150
                .Axes(xlCategory).MajorUnit = 30
                .Axes(xlCategory).MajorTickMark = xlCross
                .Axes(xlCategory).MinorTickMark = xlCross
                .Axes(xlCategory).MajorTickMark = xlCross
                
                .Axes(xlValue, xlPrimary).HasTitle = True
                .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "SUVr"
                .Axes(xlValue).Format.Line.EndArrowheadStyle = msoArrowheadStealth
                .Axes(xlValue).Format.Line.EndArrowheadLength = msoArrowheadLong
                .Axes(xlValue).Format.Line.EndArrowheadWidth = msoArrowheadWide
                .Axes(xlValue).MinimumScale = SUVr_min
                .Axes(xlValue).MaximumScale = SUVr_max
                .Axes(xlValue).MajorTickMark = xlCross

                .SetElement (msoElementPrimaryValueGridLinesNone)
                .SetElement (msoElementPrimaryCategoryGridLinesNone)
                
                .FullSeriesCollection(1).MarkerForegroundColorIndex = -4142
                .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(92, 167, 147)
                .FullSeriesCollection(1).MarkerSize = 4
                .FullSeriesCollection(1).MarkerStyle = 8
                
                .ChartArea.Font.Color = RGB(31, 78, 121)
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
            u = u + 5
        Loop

'''''''''''''''''''''''''''''''''''''Columns Autofit''''''''''''''''''''''''''''''''''''

        Range("A" & row_total + 5).Select
        Selection.Value = "Individual Graphs Cluster:"

        Range("A" & row_total + 33).Select
        Selection.Value = "Multiplots:"

        For Z = 1 To column_total_after_SUV
            Columns(Z).EntireColumn.AutoFit
        Next Z
        
'''''''''''''''''''''''''''''''''''''Top Row Formula Removal''''''''''''''''''''''''''''''''''''
        
        Rows("1:1").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

''''''''''''''''''''''''''''''''''''''Chart Lineup''''''''''''''''''''''''''''''''''''''

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
        Next
        For Chart_index = 1 To Chart_count / 3
            ActiveSheet.ChartObjects(Chart_index).Top = 620
        Next
        For Chart_index = Chart_count / 3 + 1 To ((2 * Chart_count) / 3)
            ActiveSheet.ChartObjects(Chart_index).Top = 770
        Next
        For Chart_index = ((2 * Chart_count) / 3) + 1 To Chart_count
            ActiveSheet.ChartObjects(Chart_index).Top = 920
        Next

'''''''''''''''''''''''''''''''''SUV and SUVr Multiplots'''''''''''''''''''''''''''''''''

    
    ROIs = Array(Array("SUV_W_MUBADA_202_Wbrain", "SUV_Hippocampus_L_AAL", "SUV_Hippocampus_R_AAL", "SUV_Precuneus_L_AAL", "SUV_Precuneus_R_AAL", "SUV_Putamen_L_AAL", "SUV_Putamen_R_AAL", "SUV_blcere_all") _
    , Array("SUV_W_MUBADA_202_Wbrain", "SUV_Hippocampus_L_AAL", "SUV_Hippocampus_R_AAL", "SUV_Precuneus_L_AAL", "SUV_Precuneus_R_AAL", "SUV_Putamen_L_AAL", "SUV_Putamen_R_AAL", "SUV_cerebellum_crus1_v5") _
    , Array("SUVr(blcere)_W_MUBADA_202_Wbrain", "SUVr(blcere)_Hippocampus_L_AAL", "SUVr(blcere)_Hippocampus_R_AAL", "SUVr(blcere)_Precuneus_L_AAL", "SUVr(blcere)_Precuneus_R_AAL", "SUVr(blcere)_Putamen_L_AAL", "SUVr(blcere)_Putamen_R_AAL", "SUVr(blcere)_blcere_all") _
    , Array("SUVr(Cerecrus)_W_MUBADA_202_Wbrain", "SUVr(Cerecrus)_Hippocampus_L_AAL", "SUVr(Cerecrus)_Hippocampus_R_AAL", "SUVr(Cerecrus)_Precuneus_L_AAL", "SUVr(Cerecrus)_Precuneus_R_AAL", "SUVr(Cerecrus)_Putamen_L_AAL", "SUVr(Cerecrus)_Putamen_R_AAL", "SUVr(Cerecrus)_cerebellum_crus1_v5"))

    ROIs_length = UBound(ROIs)

    For r = 0 To ROIs_length
        Sub_ROIs_length = UBound(ROIs(r))

        For rr = 0 To Sub_ROIs_length
            Range("C1").Select
            Cells.Find(What:=ROIs(r)(rr), After:=ActiveCell, LookIn:= _
                xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                xlNext, MatchCase:=False, SearchFormat:=True).Activate
            Col_array(r, rr) = Col_lett(Selection.Column)
        Next rr

        Range("B2:B" & row_total).Select
        Multi_range = Col_array(r, 0) & "2:" & Col_array(r, 0) & row_total & "," & Col_array(r, 1) & "2:" & Col_array(r, 1) & row_total & "," & Col_array(r, 2) & "2:" & Col_array(r, 2) & row_total & "," & Col_array(r, 3) & "2:" & Col_array(r, 3) & row_total & "," & Col_array(r, 4) & "2:" & Col_array(r, 4) & row_total & "," & Col_array(r, 5) & "2:" & Col_array(r, 5) & row_total & "," & Col_array(r, 6) & "2:" & Col_array(r, 6) & row_total & "," & Col_array(r, 7) & "2:" & Col_array(r, 7) & row_total
        Range(Multi_range).Select
    
        ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
        Multiplot = Sheet_Name & "!$" & "B" & "$2:$" & "B" & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 0) & "$2:$" & Col_array(r, 0) & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 1) & "$2:$" & Col_array(r, 1) & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 2) & "$2:$" & Col_array(r, 2) & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 3) & "$2:$" & Col_array(r, 3) & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 4) & "$2:$" & Col_array(r, 4) & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 5) & "$2:$" & Col_array(r, 5) & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 6) & "$2:$" & Col_array(r, 6) & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 7) & "$2:$" & Col_array(r, 7) & "$" & row_total
    
        With ActiveChart
            .SetSourceData Source:=Range(Multiplot)
            .ChartType = xlXYScatterSmooth
            .HasLegend = True
            With .SeriesCollection(1)
                .Name = "MUBADA"
                .Format.Fill.ForeColor.RGB = RGB(19, 149, 186)
                .Format.Line.ForeColor.RGB = RGB(19, 149, 186)
                .Format.Line.Weight = 1.25
                .MarkerStyle = 1
                .MarkerSize = 5
                .Format.Fill.Visible = msoFalse
                .MarkerStyle = -4168
                .Format.Line.Visible = msoTrue
                .Format.Line.DashStyle = msoLineSysDot
            End With
            With .SeriesCollection(2)
                .Name = "Hippocampus_L"
                .Format.Fill.ForeColor.RGB = RGB(13, 60, 85)
                .Format.Line.ForeColor.RGB = RGB(13, 60, 85)
                .Format.Line.Weight = 1.25
                .MarkerStyle = 1
                .MarkerSize = 5
                .Format.Fill.Visible = msoFalse
                .MarkerStyle = -4168
                .Format.Line.Visible = msoTrue
                .Format.Line.DashStyle = msoLineSysDot
            End With
            With .SeriesCollection(3)
                .Name = "Hippocampus_R"
                .Format.Fill.ForeColor.RGB = RGB(192, 46, 29)
                .Format.Line.ForeColor.RGB = RGB(192, 46, 29)
                .Format.Line.Weight = 1.25
                .MarkerStyle = 1
                .MarkerSize = 5
                .Format.Fill.Visible = msoFalse
                .MarkerStyle = -4168
                .Format.Line.Visible = msoTrue
                .Format.Line.DashStyle = msoLineSysDot
            End With
            With .SeriesCollection(4)
                .Name = "Precuneus_L"
                .Format.Fill.ForeColor.RGB = RGB(241, 108, 32)
                .Format.Line.ForeColor.RGB = RGB(241, 108, 32)
                .Format.Line.Weight = 1.25
                .MarkerStyle = 1
                .MarkerSize = 5
                .Format.Fill.Visible = msoFalse
                .MarkerStyle = -4168
                .Format.Line.Visible = msoTrue
                .Format.Line.DashStyle = msoLineSysDot
            End With
            With .SeriesCollection(5)
                .Name = "Precuneus_R"
                .Format.Fill.ForeColor.RGB = RGB(239, 139, 44)
                .Format.Line.ForeColor.RGB = RGB(239, 139, 44)
                .Format.Line.Weight = 1.25
                .MarkerStyle = 1
                .MarkerSize = 5
                .Format.Fill.Visible = msoFalse
                .MarkerStyle = -4168
                .Format.Line.Visible = msoTrue
                .Format.Line.DashStyle = msoLineSysDot
            End With
            With .SeriesCollection(6)
                .Name = "Putamen_L"
                .Format.Fill.ForeColor.RGB = RGB(235, 200, 68)
                .Format.Line.ForeColor.RGB = RGB(235, 200, 68)
                .Format.Line.Weight = 1.25
                .MarkerStyle = 1
                .MarkerSize = 5
                .Format.Fill.Visible = msoFalse
                .MarkerStyle = -4168
                .Format.Line.Visible = msoTrue
                .Format.Line.DashStyle = msoLineSysDot
            End With
            With .SeriesCollection(7)
                .Name = "Putamen_R"
                .Format.Fill.ForeColor.RGB = RGB(162, 184, 108)
                .Format.Line.ForeColor.RGB = RGB(162, 184, 108)
                .Format.Line.Weight = 1.25
                .MarkerStyle = 1
                .MarkerSize = 5
                .Format.Fill.Visible = msoFalse
                .MarkerStyle = -4168
                .Format.Line.Visible = msoTrue
                .Format.Line.DashStyle = msoLineSysDot
            End With
            With .SeriesCollection(8)
                .Name = "Cerebellum"
                .Format.Fill.ForeColor.RGB = RGB(92, 167, 147)
                .Format.Line.ForeColor.RGB = RGB(92, 167, 147)
                .Format.Line.Weight = 1.25
                .MarkerStyle = 1
                .MarkerSize = 5
                .Format.Fill.Visible = msoFalse
                .MarkerStyle = -4168
                .Format.Line.Visible = msoTrue
                .Format.Line.DashStyle = msoLineSysDot
            End With
            .ChartTitle.Text = ROIs(r)(7)
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (mins)"
            .Axes(xlCategory).MinimumScale = 0
            .Axes(xlCategory).MaximumScale = 150
            .Axes(xlValue, xlPrimary).HasTitle = True
    
            If r < 2 Then
                .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "SUV"
                .Axes(xlValue).MinimumScale = SUV_min
                .Axes(xlValue).MaximumScale = SUV_max
            Else
                .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "SUVr"
                .Axes(xlValue).MinimumScale = SUVr_min
                .Axes(xlValue).MaximumScale = SUVr_max
            End If
        End With
        
        For Each vAxis In Array(xlCategory, xlValue)
            With ActiveChart.Axes(vAxis)
                .TickLabels.Font.Name = "Consolas"
                .TickLabels.Font.Size = 10
                .MajorTickMark = xlOutside
                .MinorTickMark = xlOutside
                .Format.Line.EndArrowheadStyle = msoArrowheadTriangle
                .AxisTitle.Format.TextFrame2.TextRange.Font.Spacing = 1.4
                .AxisTitle.Format.TextFrame2.TextRange.Font.Size = 11
                With .Format.Line
                    .Visible = msoTrue
                    .ForeColor.ObjectThemeColor = msoThemeColorText1
                    .ForeColor.TintAndShade = 0
                    .ForeColor.Brightness = 0.349999994
                    .Transparency = 0
                End With
            End With
        Next vAxis
        
        Multiplot_index = ActiveSheet.ChartObjects.Count
        ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
        ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesNone)
        ActiveSheet.ChartObjects(Multiplot_index).Width = 600
        ActiveSheet.ChartObjects(Multiplot_index).Height = 350
        If r = 0 Then
            ActiveSheet.ChartObjects(Multiplot_index).Top = 1800
            ActiveSheet.ChartObjects(Multiplot_index).Left = 0
        ElseIf r = 1 Then
            ActiveSheet.ChartObjects(Multiplot_index).Top = 1800
            ActiveSheet.ChartObjects(Multiplot_index).Left = 600
        ElseIf r = 2 Then
            ActiveSheet.ChartObjects(Multiplot_index).Top = 2150
            ActiveSheet.ChartObjects(Multiplot_index).Left = 0
        ElseIf r = 3 Then
            ActiveSheet.ChartObjects(Multiplot_index).Top = 2150
            ActiveSheet.ChartObjects(Multiplot_index).Left = 600
        End If
        ActiveChart.ChartStyle = 241
    Next r
    
'''''''''''''''''''''''''''''''''''
    ROIs = Array(Array("SUV_Occipital_L", "SUV_Occipital_R", "SUV_Parietal_L", "SUV_Parietal_R", "SUV_Temporal_L", "SUV_Temporal_R", "SUV_Frontal_L", "SUV_Frontal_R", "SUV_blcere_all") _
    , Array("SUV_Occipital_L", "SUV_Occipital_R", "SUV_Parietal_L", "SUV_Parietal_R", "SUV_Temporal_L", "SUV_Temporal_R", "SUV_Frontal_L", "SUV_Frontal_R", "SUV_cerebellum_crus1_v5") _
    , Array("SUVr(blcere)_Occipital_L", "SUVr(blcere)_Occipital_R", "SUVr(blcere)_Parietal_L", "SUVr(blcere)_Parietal_R", "SUVr(blcere)_Temporal_L", "SUVr(blcere)_Temporal_R", "SUVr(blcere)_Frontal_L", "SUVr(blcere)_Frontal_R", "SUVr(blcere)_blcere_all") _
    , Array("SUVr(Cerecrus)_Occipital_L", "SUVr(Cerecrus)_Occipital_R", "SUVr(Cerecrus)_Parietal_L", "SUVr(Cerecrus)_Parietal_R", "SUVr(Cerecrus)_Temporal_L", "SUVr(Cerecrus)_Temporal_R", "SUVr(Cerecrus)_Frontal_L", "SUVr(Cerecrus)_Frontal_R", "SUVr(Cerecrus)_cerebellum_crus1_v5"))
    
    ROIs_length = UBound(ROIs)
    
    For r = 0 To ROIs_length
        Sub_ROIs_length = UBound(ROIs(r))
        
        For rr = 0 To Sub_ROIs_length
            Range("C1").Select
            Cells.Find(What:=ROIs(r)(rr), After:=ActiveCell, LookIn:= _
                xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                xlNext, MatchCase:=False, SearchFormat:=True).Activate
            Col_array(r, rr) = Col_lett(Selection.Column)
        Next rr
    
        Range("B2:B" & row_total).Select
        Multi_range = Col_array(r, 0) & "2:" & Col_array(r, 0) & row_total & "," & Col_array(r, 1) & "2:" & Col_array(r, 1) & row_total & "," & Col_array(r, 2) & "2:" & Col_array(r, 2) & row_total & "," & Col_array(r, 3) & "2:" & Col_array(r, 3) & row_total & "," & Col_array(r, 4) & "2:" & Col_array(r, 4) & row_total & "," & Col_array(r, 5) & "2:" & Col_array(r, 5) & row_total & "," & Col_array(r, 6) & "2:" & Col_array(r, 6) & row_total & "," & Col_array(r, 7) & "2:" & Col_array(r, 7) & row_total & "," & Col_array(r, 8) & "2:" & Col_array(r, 8) & row_total
        Range(Multi_range).Select
    
        ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
        Multiplot = Sheet_Name & "!$" & "B" & "$2:$" & "B" & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 0) & "$2:$" & Col_array(r, 0) & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 1) & "$2:$" & Col_array(r, 1) & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 2) & "$2:$" & Col_array(r, 2) & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 3) & "$2:$" & Col_array(r, 3) & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 4) & "$2:$" & Col_array(r, 4) & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 5) & "$2:$" & Col_array(r, 5) & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 6) & "$2:$" & Col_array(r, 6) & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 7) & "$2:$" & Col_array(r, 7) & "$" & row_total & "," _
        & Sheet_Name & "!$" & Col_array(r, 8) & "$2:$" & Col_array(r, 8) & "$" & row_total
    
        With ActiveChart
            .SetSourceData Source:=Range(Multiplot)
            .ChartType = xlXYScatterSmooth
            .HasLegend = True
            
            With .SeriesCollection(1)
                .Name = "Cerebellum"
                .Format.Fill.ForeColor.RGB = RGB(92, 167, 147)
                .Format.Line.ForeColor.RGB = RGB(92, 167, 147)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            With .SeriesCollection(2)
                .Name = "Occipital_L"
                .Format.Fill.ForeColor.RGB = RGB(19, 149, 186)
                .Format.Line.ForeColor.RGB = RGB(19, 149, 186)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            With .SeriesCollection(3)
                .Name = "Occipital_R"
                .Format.Fill.ForeColor.RGB = RGB(13, 60, 85)
                .Format.Line.ForeColor.RGB = RGB(13, 60, 85)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            With .SeriesCollection(4)
                .Name = "Parietal_L"
                .Format.Fill.ForeColor.RGB = RGB(192, 46, 29)
                .Format.Line.ForeColor.RGB = RGB(192, 46, 29)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            With .SeriesCollection(5)
                .Name = "Parietal_R"
                .Format.Fill.ForeColor.RGB = RGB(241, 108, 32)
                .Format.Line.ForeColor.RGB = RGB(241, 108, 32)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            With .SeriesCollection(6)
                .Name = "Temporal_L"
                .Format.Fill.ForeColor.RGB = RGB(239, 139, 44)
                .Format.Line.ForeColor.RGB = RGB(239, 139, 44)
                  .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            With .SeriesCollection(7)
                .Name = "Temporal_R"
                .Format.Fill.ForeColor.RGB = RGB(235, 200, 68)
                .Format.Line.ForeColor.RGB = RGB(235, 200, 68)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            With .SeriesCollection(8)
                .Name = "Frontal_L"
                .Format.Fill.ForeColor.RGB = RGB(162, 184, 108)
                .Format.Line.ForeColor.RGB = RGB(162, 184, 108)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            With .SeriesCollection(9)
                .Name = "Frontal_R"
                .Format.Fill.ForeColor.RGB = RGB(40, 82, 72)
                .Format.Line.ForeColor.RGB = RGB(40, 82, 72)
                .Format.Line.Weight = 1.25
                .MarkerStyle = -4142
                .MarkerSize = 5
            End With
            .ChartTitle.Text = ROIs(r)(8)
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (mins)"
            .Axes(xlCategory).MinimumScale = 0
            .Axes(xlCategory).MaximumScale = 150
            .Axes(xlValue, xlPrimary).HasTitle = True
            
            If r < 2 Then
                .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "SUV"
                .Axes(xlValue).MinimumScale = SUV_min
                .Axes(xlValue).MaximumScale = SUV_max
            Else
                .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "SUVr"
                .Axes(xlValue).MinimumScale = SUVr_min
                .Axes(xlValue).MaximumScale = SUVr_max
            End If
        End With
        
        Multiplot_index = ActiveSheet.ChartObjects.Count
        ActiveSheet.ChartObjects(Multiplot_index).Width = 600
        ActiveSheet.ChartObjects(Multiplot_index).Height = 350
        If r = 0 Then
            ActiveSheet.ChartObjects(Multiplot_index).Top = 2500
            ActiveSheet.ChartObjects(Multiplot_index).Left = 0
        ElseIf r = 1 Then
            ActiveSheet.ChartObjects(Multiplot_index).Top = 2500
            ActiveSheet.ChartObjects(Multiplot_index).Left = 600
        ElseIf r = 2 Then
            ActiveSheet.ChartObjects(Multiplot_index).Top = 2850
            ActiveSheet.ChartObjects(Multiplot_index).Left = 0
        ElseIf r = 3 Then
            ActiveSheet.ChartObjects(Multiplot_index).Top = 2850
            ActiveSheet.ChartObjects(Multiplot_index).Left = 600
        End If
        ActiveChart.ChartStyle = 241
    Next r
    
    Range("A70").Select
    ActiveWindow.Zoom = 80
    ActiveSheet.Name = Subject_study
    
End Sub
