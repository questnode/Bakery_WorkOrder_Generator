'   ### set basic parameters
Private work_date As String
Private week_day_num As Integer
Private pro_subfolder_path_row As Integer, pro_subfolder_path_col As Integer, prepro_subfolder_path_row As Integer, staff_subfolder_path_col As Integer, staff_subfolder_path_row As Integer, prepro_subfolder_path_col As Integer

'### mixing work list row gaps and offsets
Private mixing_line_gap As Integer, mixing_weight_offset As Integer, mixing_shift_offset As Integer, mixing_staff_offset As Integer, mixing_time_offset As Integer

'### mixing levain override cells
Private levain_override_am As Double, levain_override_pm As Double, levain_override_night As Double

'### set string tags
Private dough_tag As String, product_tag As String, type_dough As String, type_bread As String, type_starter As String, shift_indicator As String

'### set column numbers
Private read_name_col As Integer, read_qty_col As Integer, read_weight_col As Integer, read_waste_col As Integer, mixing_name_col As Integer, mixing_weight_col As Integer
Private product_name_col As Integer, product_qty_col As Integer
Private item_type_col As Integer
'### set column numbers for master list
Private master_sheet As String, master_stage_col As Integer, master_item_col As Integer, master_qty_col As Integer, master_staff_col As Integer, master_time_col As Integer, master_shift_col As Integer, master_sheetname_col As Integer, master_write_start_row As Integer, master_write_row As Integer, master_date_row As Integer, master_date_col As Integer

'### variable for names of different stages
Private starter_stage As String, dough_stage As String, scaling_stage As String, shaping_stage As String, baking_stage As String, packing_stage As String

'### variable for owners sheets
Private owner_starter_sheet As String, owner_dough_sheet As String, owner_scaling_sheet As String, owner_shaping_sheet As String, owner_baking_sheet As String

Sub GenerateNew()
    
    Application.ScreenUpdating = False

    Path = ActiveWorkbook.Path


'    ### declare worksheet names
    Dim recipe_sheet As String, control_sheet_name As String, source_sheet As String, mixing_list As String, starters_list As String, product_list As String, full_list As String, production_sheets As String, prepro_sheets As String
    control_sheet_name = "Control"
    recipe_sheet = "RecipeDB"
    master_sheet = "Master List"
    mixing_list = "Mixing Works"
    starters_list = "Starter Works"
    product_list = "Product Works"
    full_list = "All Works"
    
'    Sheets(control_sheet_name).Cells(5, 10).Replace " ", vbNullString, xlPart
'    Sheets(control_sheet_name).Cells(4, 10).Replace " ", vbNullString, xlPart
    
    pro_subfolder_path_row = 5
    pro_subfolder_path_col = 10
    prepro_subfolder_path_row = 4
    prepro_subfolder_path_col = 10
    staff_subfolder_path_row = 6
    staff_subfolder_path_col = 10
    
    production_sheets = import_files(control_sheet_name, Sheets(control_sheet_name).Cells(pro_subfolder_path_row, pro_subfolder_path_col))
    prepro_sheets = import_files(control_sheet_name, Sheets(control_sheet_name).Cells(prepro_subfolder_path_row, prepro_subfolder_path_col))
    
    Dim support_sheets As String
    Dim removal_sheets() As Variant
    
'    Workbooks.Open Filename:=ActiveWorkbook.Path & "\" & Sheets(control_sheet_name).Cells(6, 10) & "\" & Dir(FolderPath & "*.xls*"), ReadOnly:=True
'    For Each Sheet In ActiveWorkbook.Sheets
'       Sheet.Copy Before:=ThisWorkbook.Sheets(1)
'    Next Sheet

    support_sheets = import_files(control_sheet_name, Sheets(control_sheet_name).Cells(staff_subfolder_path_row, staff_subfolder_path_col))
    removal_sheets = breakup_list_string(support_sheets)
    
'    production_sheets = Sheets(control_sheet_name).Cells(5, 10)
'    prepro_sheets = Sheets(control_sheet_name).Cells(4, 10)
        
    '### set levain override qty
    Dim levain_am_override_row As Integer, levain_am_override_col As Integer, levain_pm_override_row As Integer, levain_pm_override_col As Integer, levain_night_override_row As Integer, levain_night_override_col As Integer
    
    levain_am_override_row = 7
    levain_am_override_col = 10
    levain_pm_override_row = 8
    levain_pm_override_col = 10
    levain_night_override_row = 9
    levain_night_override_col = 10
    
    levain_override_am = Sheets(control_sheet_name).Cells(levain_am_override_row, levain_am_override_col)
    levain_override_pm = Sheets(control_sheet_name).Cells(levain_pm_override_row, levain_pm_override_col)
    levain_override_night = Sheets(control_sheet_name).Cells(levain_night_override_row, levain_night_override_col)
        
        
'   ### set basic parameters
    work_date = Sheets(control_sheet_name).Cells(3, 10)
    week_day_num = Weekday(work_date, 2)
    
'   ### set checkbox names
    Dim checkbox_doughs As String, checkbox_starters As String, checkbox_processing As String, checkbox_baking As String, checkbox_packing As String
    checkbox_doughs = "CheckBox_Doughs"
    checkbox_starters = "CheckBox_Starters"
    checkbox_scaling = "CheckBox_Scaling"
    checkbox_shaping = "CheckBox_Shaping"
    checkbox_baking = "CheckBox_Baking"
    checkbox_packing = "CheckBox_Packing"
        
'   ### 0.2 ### Set and declare control variables for doing generating different stages of process
    Dim do_doughs As Boolean, do_starters As Boolean, do_scaling As Boolean, do_shaping As Boolean, do_baking As Boolean, do_packing As Boolean
    If Sheets(control_sheet_name).CheckBoxes(checkbox_doughs).Value = xlOn Then
        do_doughs = True
    End If
    If Sheets(control_sheet_name).CheckBoxes(checkbox_starters).Value = xlOn Then
        do_starters = True
    End If
    If Sheets(control_sheet_name).CheckBoxes(checkbox_scaling).Value = xlOn Then
        do_scaling = True
    End If
    If Sheets(control_sheet_name).CheckBoxes(checkbox_shaping).Value = xlOn Then
        do_shaping = True
    End If
    If Sheets(control_sheet_name).CheckBoxes(checkbox_baking).Value = xlOn Then
        do_baking = True
    End If
    If Sheets(control_sheet_name).CheckBoxes(checkbox_packing).Value = xlOn Then
        do_packing = True
    End If
    
    ' ### set sheet list array variables
    Dim sheet_output() As String, temp_sheets() As String, soaker_sheets() As String

    '### mixing work list row gaps and offsets
    mixing_line_gap = 5
    mixing_weight_offset = 1
    mixing_shift_offset = 2
    mixing_staff_offset = 3
    mixing_time_offset = 4
    
    '### set string tags
    dough_tag = "Dough, "
    product_tag = "FI, "
    type_dough = "Dough"
    type_cut = "Cutting"
    type_shape = "Shaping"
    type_bake = "Baking"
    type_starter = "Starter"
    shift_indicator = "Production Dept:"
    
    '### set column numbers
    read_name_col = 3
    read_qty_col = 9
    read_weight_col = 11
    read_waste_col = 7
    mixing_name_col = 1
    mixing_weight_col = 3

    product_name_col = 1
    product_qty_col = 3

    item_type_col = 2
    
    master_stage_col = 1
    master_item_col = 2
    master_qty_col = 3
    master_unit_col = 4
    master_staff_col = 5
    master_time_col = 6
    master_shift_col = 7
    master_sheetname_col = 8
    master_write_start_row = 3
    master_write_row = master_write_start_row
    master_date_row = 1
    master_date_col = 6
    
'   ### set names for stages
    starter_stage = "Starters"
    dough_stage = "Doughs"
    scaling_stage = "Scaling"
    shaping_stage = "Shaping"
    baking_stage = "Baking"
    packing_stage = "Packing"

    Dim staffing() As Variant
    
    Dim template_master As String
    template_master = "Template_Master List"
    
    ' ### set names for owner sheets
    owner_starter_sheet = "Owners_Starters"
    owner_dough_sheet = "Owners_Doughs"
    owner_scaling_sheet = "Owners_Scaling"
    owner_shaping_sheet = "Owners_Shaping"
    owner_baking_sheet = "Owners_Baking"
    

    If do_doughs = True Or do_starters = True Or do_processing = True Or do_baking = True Or do_packing = True Then
        Sheets(template_master).Copy Before:=Sheets(1)
        Sheets(template_master & " (2)").name = master_sheet
        Sheets(master_sheet).Cells(master_date_row, master_date_col) = work_date
    End If

    If do_doughs = True And production_sheets <> "" Then
        Call run_mixing(mixing_list, production_sheets)
        Call remove_extra(mixing_list, "starters")
        Call create_mixing_sheets(mixing_list, dough_stage)
    End If
        
    If do_starters = True And prepro_sheets <> "" Then
        Call run_mixing(starters_list, prepro_sheets)
        Call remove_extra(starters_list, "doughs")
        Call create_mixing_sheets(starters_list, starter_stage)
    End If
    
    If do_scaling = True Or do_shaping = True Or do_baking Then
        Call run_processing(production_sheets, do_scaling, do_shaping, do_baking)
        
    End If
    
    '### sort master sheet
    
    Sheets(master_sheet).Sort.SortFields.Clear
    With Sheets(master_sheet).Sort
     .SortFields.Add Key:=Range("E1"), Order:=xlAscending
     .SortFields.Add Key:=Range("F1"), Order:=xlAscending
     .SetRange Range("A2:H500")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    '### create single staff work sheet
    Dim staff_name As String, template_single_staff As String
    template_single_staff = "Template_Staff_Processing"
    
    Dim staff_start_row As Integer, current_staff_write_row As Integer
    staff_start_row = 3
    current_staff_write_row = staff_start_row
    
    Dim staff_write_stage_col As Integer, staff_write_item_col As Integer, staff_write_qty_col As Integer, staff_write_unit_col As Integer, staff_write_target_start_col As Integer, staff_write_target_finish_col As Integer
    staff_write_stage_col = 1
    staff_write_item_col = 2
    staff_write_qty_col = 3
    staff_write_unit_col = 4
    staff_write_target_start_col = 5
    staff_write_target_finish_col = 6
    staff_write_name_col = 3
    staff_write_name_row = 1
    staff_write_date_row = 1
    staff_write_date_col = 1
    
    Dim staff_list() As String
    ReDim staff_list(1 To 1)
    
    For t = 3 To 300
        If Sheets(master_sheet).Cells(t, master_stage_col) = "" Then
            Exit For
        Else
            If Sheets(master_sheet).Cells(t, master_staff_col) <> staff_name Then
                staff_name = Sheets(master_sheet).Cells(t, master_staff_col)
                Sheets(template_single_staff).Copy Before:=Sheets(1)
                Sheets(template_single_staff & " (2)").name = staff_name
                staff_list(UBound(staff_list)) = staff_name
                Sheets(staff_name).Cells(staff_write_name_row, staff_write_name_col) = staff_name
                Sheets(staff_name).Cells(staff_write_date_row, staff_write_date_col) = work_date
                ReDim Preserve staff_list(1 To (UBound(staff_list) + 1))
                current_staff_write_row = staff_start_row
            End If
            
            Sheets(staff_name).Cells(3, 1).EntireRow.Copy
            Sheets(staff_name).Cells(current_staff_write_row, 1).EntireRow.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
            
            Sheets(staff_name).Cells(current_staff_write_row, staff_write_stage_col) = Sheets(master_sheet).Cells(t, master_stage_col)
            Sheets(staff_name).Cells(current_staff_write_row, staff_write_item_col) = Sheets(master_sheet).Cells(t, master_item_col)
            If Sheets(master_sheet).Cells(t, master_stage_col) = starter_stage Or Sheets(master_sheet).Cells(t, master_stage_col) = dough_stage Then
                Sheets(staff_name).Cells(current_staff_write_row, staff_write_qty_col) = Round(Sheets(master_sheet).Cells(t, master_qty_col) * 1000, 0)
                'Sheets(staff_name).Cells(current_staff_write_row, staff_write_qty_col).NumberFormat = "#,##0"
                Sheets(staff_name).Cells(current_staff_write_row, staff_write_unit_col) = "g"
                Sheets(staff_name).Cells(current_staff_write_row, staff_write_target_finish_col) = Sheets(master_sheet).Cells(t, master_time_col)
            Else
                Sheets(staff_name).Cells(current_staff_write_row, staff_write_qty_col) = Val(Sheets(master_sheet).Cells(t, master_qty_col))
                Sheets(staff_name).Cells(current_staff_write_row, staff_write_unit_col) = "pc"
                Sheets(staff_name).Cells(current_staff_write_row, staff_write_target_start_col) = Sheets(master_sheet).Cells(t, master_time_col)
            End If
            Sheets(staff_name).Cells(current_staff_write_row, staff_write_qty_col).NumberFormat = "#,##0"
            current_staff_write_row = current_staff_write_row + 1
        
        End If
    
    
    Next t
    
    '###final output of all sheets
    Dim output_list() As Variant
   
    ReDim output_list(1 To 1)
'    test = UBound(output_list)
    
'    Dim empty_max As Integer, empty_count As Integer
'    empty_count = 0
'    empty_max = 100
    output_list(UBound(output_list)) = "Master List"
    ReDim Preserve output_list(1 To (UBound(output_list) + 1))
    
    For u = 1 To UBound(staff_list)
        output_list(UBound(output_list)) = staff_list(u)
        ReDim Preserve output_list(1 To (UBound(output_list) + 1))
    
    Next u
    ReDim Preserve output_list(1 To (UBound(output_list) - 1))
    For i = master_write_start_row To 300
        'If empty_count > empty_max Then
            
        If Sheets(master_sheet).Cells(i, master_stage_col) = "" Then
            Exit For
'        ElseIf Sheets(master_sheet).Cells(i, master_stage_col) <> "" Then
'            Sheets(master_sheet).Cells(master_write_start_row, 1).EntireRow.Copy
'                Sheets(master_sheet).Cells(i, 1).EntireRow.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'            SkipBlanks:=False, Transpose:=False
        
        ElseIf Sheets(master_sheet).Cells(i, master_sheetname_col) <> "" Then
            
            output_list(UBound(output_list)) = Sheets(master_sheet).Cells(i, master_sheetname_col)
            ReDim Preserve output_list(1 To (UBound(output_list) + 1))
            empty_count = 0
        End If
    Next i
    
    '### format master list
    For i = master_write_start_row To 300
        If Sheets(master_sheet).Cells(i, master_stage_col) = "" Then
            Exit For
        Else
            Sheets(master_sheet).Cells(master_write_start_row, 1).EntireRow.Copy
                Sheets(master_sheet).Cells(i, 1).EntireRow.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        End If
    Next i


    
    ReDim Preserve output_list(1 To (UBound(output_list) - 1))
    

'    ReDim Preserve output_list(1 To (UBound(output_list) + 1))
'    output_list(UBound(output_list)) = "Mixing Works"
'    ReDim Preserve output_list(1 To (UBound(output_list) + 1))
'    output_list(UBound(output_list)) = "Starter Works"

    
    Application.DisplayAlerts = False
    
    If production_sheets <> "" Then
        If do_starters = True Then
            Sheets(mixing_list).Delete
        End If
        Sheets(breakup_list_string(production_sheets)).Delete
    End If
    If prepro_sheets <> "" Then
        If do_starters = True Then
            Sheets(starters_list).Delete
        End If
        Sheets(breakup_list_string(prepro_sheets)).Delete
    End If
    
'    Sheets("Owners_Baking").Delete
'    Sheets("Owners_Shaping").Delete
'    Sheets("Owners_Scaling").Delete
'    Sheets("Owners_Doughs").Delete
'    Sheets("Owners_Starters").Delete
    
    Sheets(removal_sheets).Delete
    
    Application.DisplayAlerts = True
    
    For x = UBound(output_list) To LBound(output_list) Step -1
        Worksheets(output_list(x)).Move Before:=Worksheets(1)
    Next x
    
    Sheets(master_sheet).Sort.SortFields.Clear
    With Sheets(master_sheet).Sort
     .SortFields.Add Key:=Range("F1"), Order:=xlAscending
     .SortFields.Add Key:=Range("A1"), Order:=xlAscending
     .SetRange Range("A2:H500")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    For r = 2 To 600
        
        If Sheets(master_sheet).Cells(r, master_stage_col) = "" Then
            Exit For
        ElseIf Sheets(master_sheet).Cells(r, master_stage_col) = dough_stage Or Sheets(master_sheet).Cells(r, master_stage_col) = starter_stage Then
            Sheets(master_sheet).Cells(r, master_unit_col) = "g"
            Sheets(master_sheet).Cells(r, master_qty_col) = Round(Sheets(master_sheet).Cells(r, master_qty_col) * 1000, 0)
            'Sheets(master_sheet).Cells(r, master_qty_col).NumberFormat = "#,##0"
        Else
            Sheets(master_sheet).Cells(r, master_unit_col) = "pc"
        End If
        
        Sheets(master_sheet).Cells(r, master_qty_col).NumberFormat = "#,##0"
    Next r
    
    Sheets(master_sheet).Columns("H:H").Delete shift:=xlToLeft
    
    Sheets(output_list).Move
    
    Application.ScreenUpdating = True
End Sub
Function import_files(read_sheet As String, subfolder As String) As String

    Dim FolderPath As String
    ', subfolder_pro As String, subfolder_prepro As String
    Dim Filename As String, all_names As String
    Dim Sheet As Worksheet

    FolderPath = ActiveWorkbook.Path & "\" & subfolder & "\"
    
    Filename = Dir(FolderPath & "*.xls*")
    
    Do While Filename <> ""
     Workbooks.Open Filename:=FolderPath & Filename, ReadOnly:=True
     For Each Sheet In ActiveWorkbook.Sheets
        Sheet.Copy Before:=ThisWorkbook.Sheets(1)
        If subfolder <> Sheets(read_sheet).Cells(staff_subfolder_path_row, staff_subfolder_path_col) Then
            Sheets(1).name = Left(Filename, InStr(Filename, ".") - 1)
            'ThisWorkbook.Sheets(1).name = Left(Filename, InStr(Filename, ".") - 1)
        End If
        all_names = all_names & "," & ThisWorkbook.Sheets(1).name
        'all_names = all_names & "," & Left(Filename, InStr(Filename, ".") - 1)
     Next Sheet
     Workbooks(Filename).Close savechanges:=False
     Filename = Dir()
    Loop
    
    If all_names <> "" Then
        all_names = Right(all_names, Len(all_names) - 1)
    End If
    
    
    import_files = all_names

End Function

Function breakup_list_string(list_string As String) As Variant

    Dim list_array() As Variant
    ReDim list_array(1 To 1)
    
    list_array(1) = list_string
    
    For i = 2 To 11
        If InStr(list_array(i - 1), ",") > 0 Then
            ReDim Preserve list_array(1 To i)
            list_array(i) = Right(list_array(i - 1), Len(list_array(i - 1)) - InStr(list_array(i - 1), ","))
            list_array(i - 1) = Left(list_array(i - 1), InStr(list_array(i - 1), ",") - 1)
        Else
            ReDim Preserve list_array(1 To i - 1)
            Exit For
        End If
    Next i

    breakup_list_string = list_array()

End Function
Sub run_mixing(work_list As String, read_list As String)
    
    
    Dim production_sheets() As Variant
    
    Dim recipe_sheet As String
    recipe_sheet = "RecipeDB"
    
    '### running variables
    Dim empty_max As Integer, empty_count As Integer, mixing_write_row As Integer, product_write_row As Integer, full_list_write_row As Integer
    Dim shift As String
    
    empty_max = 30

    mixing_write_row = 1
    product_write_row = 1
    full_list_write_row = 1
    
    '### seperate sheets names into arrays
    production_sheets = breakup_list_string(read_list)
    
    '### create working lists of items
    Sheets.Add(Before:=Sheets(1)).name = work_list

    For s = 1 To UBound(production_sheets)
    
        source_sheet = production_sheets(s)
        empty_count = 0
        
        For i = 1 To 50
         If Sheets(source_sheet).Cells(i, 2) = shift_indicator Then
            shift = Sheets(source_sheet).Cells(i, 5)
            Exit For
         End If
        Next i
        
        empty_count = 0
        
        For i = 1 To 500
    
            If InStr(Sheets(source_sheet).Cells(i, read_name_col), dough_tag) > 0 Then
                Sheets(work_list).Cells(mixing_write_row, mixing_name_col) = Right(Sheets(source_sheet).Cells(i, read_name_col), Len(Sheets(source_sheet).Cells(i, read_name_col)) - InStr(Sheets(source_sheet).Cells(i, read_name_col), dough_tag) + 1)
                For w = i To i + 30
                    If Sheets(source_sheet).Cells(w, read_name_col) = "Total To Produce:" Then
                        Sheets(work_list).Cells(mixing_write_row + mixing_weight_offset, mixing_name_col) = Sheets(source_sheet).Cells(w, read_weight_col)
                        Exit For
                    End If
                Next w
                Sheets(work_list).Cells(mixing_write_row + mixing_shift_offset, mixing_name_col) = shift
                mixing_write_row = mixing_write_row + mixing_line_gap
                empty_count = 0
                
            ElseIf empty_count > empty_max Then
                Exit For
            End If
    
            empty_count = empty_count + 1
    
        Next i
        
    Next s


'   ### Run recipes
    Dim recipe As Variant
    Dim dough_weight As Double
    Dim insert_range As String, mix_name As String
    empty_count = 0
    
    Dim levain_name As String
    levain_name = "Starter, Levain"

    For s = 1 To 100
        If Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap), mixing_name_col) <> "" Then
            
            mix_name = Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap), mixing_name_col)
            
            If InStr(mix_name, "[") > 0 Then
                mix_name = Left(mix_name, InStr(mix_name, "[") - 2)
                Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap), mixing_name_col) = mix_name
            End If
            
                recipe = get_recipe(mix_name, recipe_sheet, Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap) + 1, mixing_name_col), 0)
            
            empty_count = 0
            
            For d = 1 To (UBound(recipe, 2) - 1)
                Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap), mixing_name_col + d) = recipe(1, d)
                Sheets(work_list).Cells(2 + ((s - 1) * mixing_line_gap), mixing_name_col + d) = recipe(2, d)
                Sheets(work_list).Cells(3 + ((s - 1) * mixing_line_gap), mixing_name_col + d) = recipe(3, d)
            
                If (InStr(recipe(1, d), dough_tag) > 0 Or InStr(recipe(1, d), "Starter, ") > 0 Or InStr(recipe(1, d), "Soaker, ") > 0) And (InStr(recipe(1, d), levain_name & " [") = 0 Or InStr(mix_name, "Starter, ") = 0) Then
                    insert_range = CStr((1 + ((s) * mixing_line_gap))) + ":" + CStr(((s) * mixing_line_gap) + mixing_line_gap)
                    Sheets(work_list).Rows(insert_range).Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    Sheets(work_list).Cells((1 + ((s) * mixing_line_gap)), mixing_name_col) = recipe(1, d)
                    Sheets(work_list).Cells((1 + mixing_weight_offset + ((s) * mixing_line_gap)), mixing_name_col) = recipe(3, d)
                    Sheets(work_list).Cells((1 + mixing_shift_offset + ((s) * mixing_line_gap)), mixing_name_col) = Sheets(work_list).Cells(1 + mixing_shift_offset + ((s - 1) * mixing_line_gap), mixing_name_col)
                End If
            Next d
            
        ElseIf Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap)) = "" Then
            empty_count = empty_count + 1
        ElseIf empty_count > empty_max Then
            Exit For
        End If
    Next s

    '### merge identical mixes within same shifts
    Dim match_dough As String
    Dim empty_count2 As Integer
    empty_count = 0
    empty_count2 = 0

    For s = 1 To 200
        If Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap), mixing_name_col) <> "" Then
            match_dough = Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap), mixing_name_col)
            For m = (s + 1) To 200
                If Sheets(work_list).Cells(1 + ((m - 1) * mixing_line_gap), mixing_name_col) = match_dough And Sheets(work_list).Cells(1 + mixing_shift_offset + ((s - 1) * mixing_line_gap), mixing_name_col) = Sheets(work_list).Cells(1 + mixing_shift_offset + ((m - 1) * mixing_line_gap), mixing_name_col) Then
                    Sheets(work_list).Cells(2 + ((s - 1) * mixing_line_gap), 1) = Sheets(work_list).Cells(2 + ((s - 1) * mixing_line_gap), 1) + Sheets(work_list).Cells(2 + ((m - 1) * mixing_line_gap), 1)
                    For i = 2 To 20
                        If Sheets(work_list).Cells(1 + ((m - 1) * mixing_line_gap), i) <> "" Then
                            Sheets(work_list).Cells(3 + ((s - 1) * mixing_line_gap), i) = Sheets(work_list).Cells(3 + ((s - 1) * mixing_line_gap), i) + Sheets(work_list).Cells(3 + ((m - 1) * mixing_line_gap), i)
                        Else
                            Exit For
                        End If
                    Next i

                    insert_range = CStr((1 + ((m - 1) * mixing_line_gap))) + ":" + CStr(((m - 1) * mixing_line_gap) + mixing_line_gap)
                    Sheets(work_list).Rows(insert_range).Delete shift:=xlUp
                    m = m - 1

                ElseIf empty_count2 > empty_max Then
                    Exit For
                End If
            Next m
        Else
            Exit For
        End If
    Next s

    '### deal with levain feedings
    Dim levain_schedule_sheet As String, week_day As String
    Dim weekday_gap As Integer, recipe_row_mod As Integer, levain_count As Integer, buffer_am_col As Integer, buffer_afternoon_col As Integer, buffer_night_col As Integer, buffer_col As Integer
    Dim levain_override As Double
    'levain_name = "Starter, Levain"
    levain_schedule_sheet = "Levain_Schedule"
    weekday_gap = 5
    recipe_row_mod = 0
    
    buffer_am_col = 7
    buffer_afternoon_col = 8
    buffer_night_col = 9

    If week_day_num = 1 Then
        week_day = "Monday"
    ElseIf week_day_num = 2 Then
        week_day = "Tuesday"
    ElseIf week_day_num = 3 Then
        week_day = "Wednesday"
    ElseIf week_day_num = 4 Then
        week_day = "Thursday"
    ElseIf week_day_num = 5 Then
        week_day = "Friday"
    ElseIf week_day_num = 6 Then
        week_day = "Saturday"
    ElseIf week_day_num = 7 Then
        week_day = "Sunday"
    End If

    empty_count = 0
    For d = 1 To 200
        'temp_day = Sheets(levain_schedule_sheet).Cells(d, 1)
        If Sheets(levain_schedule_sheet).Cells(d, 1) = week_day Then
            recipe_row_mod = d - 1
            Exit For
        ElseIf empty_count > empty_max + 50 Or Sheets(levain_schedule_sheet).Cells(d, 1) = "End of Doughs" Then
            Exit For
        Else
            empty_count = empty_count + 1
        End If
    Next d

    

    empty_count = 0
    For s = 1 To 600
        mix_name = Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap), mixing_name_col)
        If s = 30 Then
            s = s
        End If
        If mix_name <> "" And mix_name = levain_name Then
            shift = Sheets(work_list).Cells(1 + mixing_shift_offset + ((s - 1) * mixing_line_gap), mixing_name_col)
            
            If shift = "AM" Then
                buffer_col = buffer_am_col
                levain_override = levain_override_am / 1000
            ElseIf shift = "Afternoon" Then
                buffer_col = buffer_afternoon_col
                levain_override = levain_override_pm / 1000
            ElseIf shift = "Night" Then
                buffer_col = buffer_night_col
                levain_override = levain_override_night / 1000
            End If
            
            mix_name = mix_name & ", " & shift
            
            If levain_override > 0 Then
                Sheets(work_list).Cells(1 + mixing_weight_offset + ((s - 1) * mixing_line_gap), mixing_name_col) = levain_override
            Else
                Sheets(work_list).Cells(1 + mixing_weight_offset + ((s - 1) * mixing_line_gap), mixing_name_col) = Sheets(work_list).Cells(1 + mixing_weight_offset + ((s - 1) * mixing_line_gap), mixing_name_col) + Sheets(levain_schedule_sheet).Cells(recipe_row_mod + 1, buffer_col)
            End If
            
            For m = recipe_row_mod + 1 To recipe_row_mod + 20
                If InStr(Sheets(levain_schedule_sheet).Cells(m, 1), mix_name) > 0 And Sheets(levain_schedule_sheet).Cells(m, 1) <> mix_name Then
                    mix_name = Sheets(levain_schedule_sheet).Cells(m, 1)
                    Exit For
                End If
            Next m
            
            Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap), mixing_name_col) = mix_name
            
            levain_count = Right(mix_name, 1) - 1
            
            recipe = get_recipe(mix_name, levain_schedule_sheet, Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap) + 1, mixing_name_col), recipe_row_mod)
            
            For d = 1 To (UBound(recipe, 2) - 1)
                
                Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap), mixing_name_col + d) = recipe(1, d)
                Sheets(work_list).Cells(2 + ((s - 1) * mixing_line_gap), mixing_name_col + d) = recipe(2, d)
                Sheets(work_list).Cells(3 + ((s - 1) * mixing_line_gap), mixing_name_col + d) = recipe(3, d)
                
                If recipe(1, d) = levain_name And levain_count > 0 Then
                    recipe(1, d) = recipe(1, d) & ", " & shift & ", Feed #" & levain_count
                    Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap), mixing_name_col + d) = recipe(1, d)
                    insert_range = CStr((1 + ((s) * mixing_line_gap))) + ":" + CStr(((s) * mixing_line_gap) + mixing_line_gap)
                    Sheets(work_list).Rows(insert_range).Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    Sheets(work_list).Cells((1 + ((s) * mixing_line_gap)), mixing_name_col) = recipe(1, d)
                    Sheets(work_list).Cells((1 + mixing_weight_offset + ((s) * mixing_line_gap)), mixing_name_col) = recipe(3, d)
                    Sheets(work_list).Cells((1 + mixing_shift_offset + ((s) * mixing_line_gap)), mixing_name_col) = Sheets(work_list).Cells(1 + mixing_shift_offset + ((s - 1) * mixing_line_gap), mixing_name_col)
                End If
            Next d
        
        ElseIf mix_name <> "" And mix_name <> levain_name And InStr(mix_name, levain_name) > 0 And InStr(mix_name, "Feed #") > 0 Then
            recipe = get_recipe(mix_name, levain_schedule_sheet, Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap) + 1, mixing_name_col), recipe_row_mod)
            levain_count = Right(mix_name, 1) - 1
            
            For d = 1 To (UBound(recipe, 2) - 1)
            
            Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap), mixing_name_col + d) = recipe(1, d)
            Sheets(work_list).Cells(2 + ((s - 1) * mixing_line_gap), mixing_name_col + d) = recipe(2, d)
            Sheets(work_list).Cells(3 + ((s - 1) * mixing_line_gap), mixing_name_col + d) = recipe(3, d)
            
                If recipe(1, d) = levain_name And levain_count > 0 Then
                    recipe(1, d) = recipe(1, d) & ", " & shift & ", Feed #" & levain_count
                    Sheets(work_list).Cells(1 + ((s - 1) * mixing_line_gap), mixing_name_col + d) = recipe(1, d)
                    insert_range = CStr((1 + ((s) * mixing_line_gap))) + ":" + CStr(((s) * mixing_line_gap) + mixing_line_gap)
                    Sheets(work_list).Rows(insert_range).Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    Sheets(work_list).Cells((1 + ((s) * mixing_line_gap)), mixing_name_col) = recipe(1, d)
                    Sheets(work_list).Cells((1 + mixing_weight_offset + ((s) * mixing_line_gap)), mixing_name_col) = recipe(3, d)
                    Sheets(work_list).Cells((1 + mixing_shift_offset + ((s) * mixing_line_gap)), mixing_name_col) = Sheets(work_list).Cells(1 + mixing_shift_offset + ((s - 1) * mixing_line_gap), mixing_name_col)
                End If
            Next d
'        ElseIf empty_count > empty_max Then
'            Exit For
'        Else
'            empty_count = empty_count + 1
        ElseIf mix_name = "" Then
            Exit For
        End If
    Next s

    'Sheets(production_sheets).Delete

End Sub
Function assign_staff(item As String, stage As String, shift As String) As Variant

    Dim owner_sheet As String
    Dim staffing(1 To 2) As Variant
    Dim empty_count As Integer, empty_max As Integer, item_name_col As Integer, staff_name_col As Integer, week_day_col_gaps As Integer, staff_start_col As Integer, target_time_start_col As Integer
    
'    owner_starter_sheet = "Owners_Starters"
'    owner_dough_sheet = "Owners_Doughs"
'    owner_processing_sheet = "Owners_Processing"
'    owner_baking_sheet = "Owners_Baking"
    
    If stage = starter_stage Then
        owner_sheet = owner_starter_sheet
    ElseIf stage = dough_stage Then
        owner_sheet = owner_dough_sheet
    ElseIf stage = scaling_stage Then
        owner_sheet = owner_scaling_sheet
    ElseIf stage = shaping_stage Then
        owner_sheet = owner_shaping_sheet
    ElseIf stage = baking_stage Then
        owner_sheet = owner_baking_sheet
    End If
    
    empty_count = 0
    empty_max = 200
    week_day_col_gaps = 2
    item_name_col = 1
    staff_start_col = 3
    target_time_start_col = 4
        
    For i = 2 To 200
        If Sheets(owner_sheet).Cells(i, item_name_col) = item Then
            staffing(1) = Sheets(owner_sheet).Cells(i, staff_start_col + (week_day_col_gaps * (week_day_num - 1)))
            staffing(2) = Sheets(owner_sheet).Cells(i, target_time_start_col + (week_day_col_gaps * (week_day_num - 1)))
            Exit For
        ElseIf Sheets(owner_sheet).Cells(i, item_name_col) = "" Then
            Exit For
        End If
    Next i

    If staffing(1) = "" Then
        staffing(1) = "Unassigned"
    End If

    assign_staff = staffing()

End Function
Sub create_mixing_sheets(read_list As String, stage As String)

    Dim staffing() As Variant

    For s = 0 To 200
        If Sheets(read_list).Cells((1 + ((s) * mixing_line_gap)), 1) = "" Then
            Exit For
        Else
            Sheets(master_sheet).Cells(master_write_row, master_stage_col) = stage
            Sheets(master_sheet).Cells(master_write_row, master_item_col) = Sheets(read_list).Cells((1 + ((s) * mixing_line_gap)), 1)
            Sheets(master_sheet).Cells(master_write_row, master_qty_col) = Sheets(read_list).Cells((1 + mixing_weight_offset + ((s) * mixing_line_gap)), 1)
            staffing = assign_staff(Sheets(master_sheet).Cells(master_write_row, master_item_col), stage, Sheets(read_list).Cells((1 + mixing_shift_offset + ((s) * mixing_line_gap)), 1))
            
            Sheets(master_sheet).Cells(master_write_row, master_staff_col) = staffing(1)
            
            Sheets(master_sheet).Cells(master_write_row, master_time_col) = staffing(2)
            Sheets(master_sheet).Cells(master_write_row, master_time_col).NumberFormat = "[$-en-US]h:mm AM/PM;@"
            
            Sheets(master_sheet).Cells(master_write_row, master_shift_col) = Sheets(read_list).Cells((1 + mixing_shift_offset + ((s) * mixing_line_gap)), 1)
            temp = create_single_mixing_sheet(read_list, 1 + ((s) * mixing_line_gap), stage, s + 1, staffing(1), staffing(2))
            Sheets(master_sheet).Cells(master_write_row, master_sheetname_col) = temp
            master_write_row = master_write_row + 1

        End If
    Next s

End Sub


Function get_recipe(dough_name As String, recipe_sheet As String, dough_weight As Double, recipe_row_mod As Integer) As Variant
    
    '### find recipe row
    Dim recipe_row, recipe_col As Integer
    recipe_row = 1
    recipe_col = 1

    For r = 1 + recipe_row_mod To 200 + recipe_row_mod
        If Sheets(recipe_sheet).Cells(r, recipe_col) = dough_name Then
            recipe_row = r
            Exit For
        ElseIf Sheets(recipe_sheet).Cells(r, recipe_col) = "End of Doughs" Then
            recipe_row = 404
            Exit For
        End If
    Next r

    '### create recipe
    Dim recipe() As Variant
    ReDim recipe(1 To 3, 1)
    Dim ingredient_read_col As Integer
    Dim ingredient_total_percent As Double
    
    ingredient_read_col = 2
    
    If recipe_row <> 404 Then
        While Sheets(recipe_sheet).Cells(1, ingredient_read_col) <> "End of Ingredients" And ingredient_read_col < 500
            If Sheets(recipe_sheet).Cells(recipe_row, ingredient_read_col) <> "" Then
                recipe(1, UBound(recipe, 2)) = Sheets(recipe_sheet).Cells(1, ingredient_read_col)
                recipe(2, UBound(recipe, 2)) = Sheets(recipe_sheet).Cells(recipe_row, ingredient_read_col)
                ingredient_total_percent = ingredient_total_percent + recipe(2, UBound(recipe, 2))
                ReDim Preserve recipe(1 To 3, UBound(recipe, 2) + 1)
            End If
            ingredient_read_col = ingredient_read_col + 1
        Wend
        
        recipe(1, UBound(recipe, 2)) = "Total"
        recipe(2, UBound(recipe, 2)) = ingredient_total_percent
        recipe(3, UBound(recipe, 2)) = dough_weight
    End If
    
    For i = 1 To (UBound(recipe, 2) - 1)
        recipe(3, i) = (dough_weight / recipe(2, UBound(recipe, 2))) * recipe(2, i)
    Next i
    
    get_recipe = recipe()


End Function
Sub remove_extra(read_list As String, removal_group As String)

        For s = 0 To 200
            If Sheets(read_list).Cells((1 + ((s) * mixing_line_gap)), 1) = "" Then
                Exit For
            ElseIf InStr(Sheets(read_list).Cells((1 + ((s) * mixing_line_gap)), 1), dough_tag) = 1 And removal_group = "doughs" Then
                mod_range = CStr((1 + ((s) * mixing_line_gap))) + ":" + CStr(((s) * mixing_line_gap) + mixing_line_gap)
                Sheets(read_list).Rows(mod_range).Delete shift:=xlUp
                s = s - 1
            ElseIf InStr(Sheets(read_list).Cells((1 + ((s) * mixing_line_gap)), 1), dough_tag) <> 1 And removal_group = "starters" Then
                mod_range = CStr((1 + ((s) * mixing_line_gap))) + ":" + CStr(((s) * mixing_line_gap) + mixing_line_gap)
                Sheets(read_list).Rows(mod_range).Delete shift:=xlUp
                s = s - 1
            End If
        Next s
    

End Sub

Function create_single_mixing_sheet(read_sheet As String, item_row As Integer, stage As String, id As Integer, staffing As Variant, time As Variant) As String

    Dim starter_template As String, dough_template As String, template As String, sheet_prefix As String
    starter_template = "Template_Starters"
    dough_template = "Template_Doughs"

    Dim name_row As Integer, name_col As Integer, qty_row As Integer, qty_col As Integer, date_row As Integer, date_col As Integer, staffing_row As Integer, staffing_col As Integer, time_row As Integer, time_col As Integer
    Dim recipe_start_row As Integer, ingredient_name_col As Integer, ingredient_qty_col As Integer, ingredient_percent_col As Integer
    
    date_row = 3
    date_col = 3
    name_row = 4
    name_col = 3
    qty_row = 5
    qty_col = 3
    staffing_row = 6
    staffing_col = 3
    time_row = 6
    time_col = 7
    shift_row = 3
    shift_col = 7
    
    recipe_start_row = 13
    ingredient_name_col = 2
    ingredient_qty_col = 5
    ingredient_percent_col = 9

    If stage = dough_stage Then
        template = dough_template
        sheet_prefix = dough_stage
    ElseIf stage = starter_stage Then
        template = starter_template
        sheet_prefix = starter_stage
    End If
    
    Sheets(template).Copy Before:=Sheets(1)
    Sheets(template & " (2)").name = sheet_prefix & " " & id
    
    Sheets(1).Cells(date_row, date_col) = work_date
    Sheets(1).Cells(shift_row, shift_col) = Sheets(read_sheet).Cells(item_row + mixing_shift_offset, 1)
    Sheets(1).Cells(name_row, name_col) = Sheets(read_sheet).Cells(item_row, 1)
    Sheets(1).Cells(qty_row, qty_col) = Sheets(read_sheet).Cells(item_row + mixing_weight_offset, 1) * 1000
    Sheets(1).Cells(staffing_row, staffing_col) = staffing
    Sheets(1).Cells(time_row, time_col) = time
    
    For s = 2 To 30
        If Sheets(read_sheet).Cells(item_row, s) <> "" Then
            Sheets(1).Cells(recipe_start_row + s - 2, ingredient_name_col) = Sheets(read_sheet).Cells(item_row, s)
            Sheets(1).Cells(recipe_start_row + s - 2, ingredient_qty_col) = Sheets(read_sheet).Cells(item_row + 2, s) * 1000
            Sheets(1).Cells(recipe_start_row + s - 2, ingredient_percent_col) = Sheets(read_sheet).Cells(item_row + 1, s)
        Else
            Exit For
        End If
    Next s
    
    'Sheets.Add(Before:=Sheets(1)).Name = sheet_prefix & " " & id
    
    create_single_mixing_sheet = sheet_prefix & " " & id

End Function
Sub run_processing(read_list As String, do_scaling As Boolean, do_shaping As Boolean, do_baking As Boolean)

    Dim source_sheet As String, read_name As String, shift As String

    Dim production_sheets() As Variant, staffing() As Variant
    production_sheets = breakup_list_string(read_list)
    
    Dim item_name_col As Integer
    item_name_col = 3
    
'    Dim owners_scaling As String, owners_shaping As String, owners_baking As String
'    owners_scaling = "Owners_Scaling"
'    owners_shaping = "Owners_Shaping"
'    owners_baking = "Owners_Baking"
    
    Dim staff As String
    Dim target_time As Double

    For s = 1 To UBound(production_sheets)
   
        source_sheet = production_sheets(s)
        
        For i = 1 To 50
            If Sheets(source_sheet).Cells(i, 2) = shift_indicator Then
                shift = Sheets(source_sheet).Cells(i, 5)
                Exit For
            End If
        Next i

        For i = 1 To 300
            read_name = Sheets(source_sheet).Cells(i, item_name_col)
            
            If do_scaling = True And InStr(read_name, "FI") > 0 Then
                staffing = assign_staff(Right(read_name, Len(read_name) - 5), scaling_stage, shift)
                staff = staffing(1)
                target_time = staffing(2)
                Call processing_write_to_master(source_sheet, read_name, scaling_stage, shift, staff, target_time, i)
            End If
            If do_shaping = True And InStr(read_name, "FI") > 0 Then
                staffing = assign_staff(Right(read_name, Len(read_name) - 5), shaping_stage, shift)
                staff = staffing(1)
                target_time = staffing(2)
                Call processing_write_to_master(source_sheet, read_name, shaping_stage, shift, staff, target_time, i)
            End If
            If do_baking = True And InStr(read_name, "FI") > 0 Then
                staffing = assign_staff(Right(read_name, Len(read_name) - 5), baking_stage, shift)
                staff = staffing(1)
                target_time = staffing(2)
                Call processing_write_to_master(source_sheet, read_name, baking_stage, shift, staff, target_time, i)
            End If
        Next i
    Next s
    
End Sub

Sub processing_write_to_master(source_sheet As String, read_name As String, stage As String, shift As String, staff As String, target_time As Double, i As Variant)

    Dim item_qty_col As Integer
    item_qty_col = 9

    Sheets(master_sheet).Cells(master_write_row, master_item_col) = Right(read_name, Len(read_name) - 5)
    Sheets(master_sheet).Cells(master_write_row, master_qty_col) = Val(Replace(Sheets(source_sheet).Cells(i, item_qty_col), ",", ""))
    Sheets(master_sheet).Cells(master_write_row, master_shift_col) = shift
    Sheets(master_sheet).Cells(master_write_row, master_stage_col) = stage
    Sheets(master_sheet).Cells(master_write_row, master_staff_col) = staff
    Sheets(master_sheet).Cells(master_write_row, master_time_col) = target_time
    Sheets(master_sheet).Cells(master_write_row, master_time_col).NumberFormat = "[$-en-US]h:mm AM/PM;@"
    master_write_row = master_write_row + 1

End Sub
Sub Generate()

    'Application.ScreenUpdating = False

'   ### 0.1 ### declare and set general control variables for various work orders
    Dim do_mixing, do_starters, do_processing, do_baking, do_packing As Boolean
    Dim control_sheet_name, checkbox_mixing, checkbox_starters, checkbox_processing, checkbox_baking, checkbox_packing, source_sheet As String
    Dim recipe_sheet As String, sheet_output() As String, temp_sheets() As String, soaker_sheets() As String
    Dim work_date As String
    Dim ingredient_fill_start_row As Integer
    control_sheet_name = "Control"
    checkbox_mixing = "CheckBox_Mixing"
    checkbox_starters = "CheckBox_Starters"
    checkbox_processing = "CheckBox_Processing"
    checkbox_baking = "CheckBox_Baking"
    checkbox_packing = "CheckBox_Packing"
    recipe_sheet = "RecipeDB"
    source_sheet = Sheets(control_sheet_name).Cells(4, 10)
    work_date = Sheets(source_sheet).Cells(5, 5)
    
    ReDim Preserve sheet_output(1 To 1)
    ReDim Preserve temp_sheets(1 To 1)
    ReDim Preserve soaker_sheets(1 To 1)
    
    ' ### 0.1.1 ###variables for processing, baking, packing
    Dim bread_tag As String, processing_template_name As String, bread_name As String, bread_dough As String, bread_sheets() As String, baking_sheets() As String
    Dim no_bread_count As Integer, bread_name_read_row As Integer, bread_name_read_col As Integer, bread_qty As Integer
    Dim first_bread As Boolean
    Dim bread_weight As Double
    ' ### end 0.1.1 ##
    
'   ## 0.1 end ###

'   ### 0.2 ### Set and declare control variables for doing mixing
    If Sheets(control_sheet_name).CheckBoxes(checkbox_mixing).Value = xlOn Then
        do_mixing = True
        Dim mixing_sheets() As String
    End If
'   ### 0.2 end ###

'   ### 0.3 ### Set and declare control variables for doing starters
    If Sheets(control_sheet_name).CheckBoxes(checkbox_starters).Value = xlOn Then
        do_starters = True
        Dim starter_name() As String, levain_shift As String
        Dim starter_weight() As Double, starter_percent_total, starter_percent_feed As Double, levain60_buffer As Double, levain20_buffer As Double
        Dim starter_count, levain_feeding_max, levain_create_count As Integer
        Dim starter_exist As Boolean
        levain_feeding_max = Sheets(control_sheet_name).Cells(5, 10)
        levain60_buffer = Sheets(control_sheet_name).Cells(6, 10)
        'levain20_buffer = Sheets(control_sheet_name).Cells(7, 10)
        levain_shift = Sheets(control_sheet_name).Cells(7, 10)
    End If
'   ### 0.3 end ###

'   ### 0.4 ### Set and declare control variables for doing processing
    If Sheets(control_sheet_name).CheckBoxes(checkbox_processing).Value = xlOn Then
        do_processing = True
    End If
'   ### 0.4 end ###
    
'   ### 0.5 ### Set and declare control variables for doing baking
    If Sheets(control_sheet_name).CheckBoxes(checkbox_baking).Value = xlOn Then
        do_baking = True
    End If
'   ### 0.5 end ###

'   ### 0.6 ### Set and declare control variables for doing packing
    If Sheets(control_sheet_name).CheckBoxes(checkbox_packing).Value = xlOn Then
        do_packing = True
    End If
'   ### 0.6 end ###


'   ### 1.1 ### generate mixing work orders
    If do_mixing = True Or do_starters = True Then
        Dim dough_name() As String
        Dim base_dough_name() As String
        Dim empty_row, dough_tag, new_sheet_tag As String, temp_dough_sheet As String, template_name As String
        Dim dough_weight() As Double
        Dim base_dough_weight() As Double
        Dim dough_name_read_col, dough_name_read_row, dough_weight_read_col, no_dough_name_count, search_max As Integer
        Dim ingredient_name_fill_col As Integer
        Dim ingredient_name_fill_row As Integer
        Dim ingredient_percent_fill_col As Integer
        Dim ingredient_percent_fill_row As Integer
        Dim ingredient_weight_fill_col As Integer
        Dim ingredient_weight_fill_row As Integer, recipe_row As Integer
        Dim first_dough, inside_dough_found As Boolean
        Dim base_dough As Variant, seek_terms() As Variant
        Dim insert_base_dough_offset As Integer
        
        dough_name_read_col = 3
        dough_name_read_row = 11
        dough_weight_read_col = 11
        no_dough_name_count = 0
        search_max = 200
        dough_tag = "Dough, "
        dough_weight_header = "Total To Produce:"
        first_dough = True
        temp_dough_sheet = "TempDoughSheet"
        template_name = "Template_Mixing"

        
'        ### 1.1.1 ### search detailed production list and compile dough names and dough weight into array
        While no_dough_name_count < search_max
        
'           ### comment ### find rows with the dough tag "Dough, ", then assign the dough name to the dough_name array
            If InStr(Sheets(source_sheet).Cells(dough_name_read_row, dough_name_read_col), dough_tag) > 0 Then
                If first_dough = True Then
                    ReDim Preserve dough_name(1 To 1)
                    ReDim Preserve dough_weight(1 To 1)
                    first_dough = False
                Else
                    ReDim Preserve dough_name(1 To (UBound(dough_name) + 1))
                    ReDim Preserve dough_weight(1 To (UBound(dough_weight) + 1))
                End If
               
                dough_name(UBound(dough_name)) = Sheets(source_sheet).Cells(dough_name_read_row, dough_name_read_col)
                
                ' ### comment ### find dough weight by looking for the row that says "Total To Produce:"
                For i = dough_name_read_row To (dough_name_read_row + search_max)
                    If Sheets(source_sheet).Cells(i, dough_name_read_col) = dough_weight_header Then
                        dough_weight(UBound(dough_weight)) = Sheets(source_sheet).Cells(i, dough_weight_read_col) * 1000
                        Exit For
                    End If
                Next i
                no_dough_name_count = 0
            Else
                no_dough_name_count = no_dough_name_count + 1
            End If
            
            dough_name_read_row = dough_name_read_row + 1
        
        Wend
        
'       ### 1.1.2 ### create new sheet (temporary) and insert dough name and weight into sheet
        Sheets.Add(Before:=Sheets(1)).name = temp_dough_sheet
        For d = 1 To (UBound(dough_name))
            Sheets(temp_dough_sheet).Cells(d, dough_name_read_col) = dough_name(d)
            Sheets(temp_dough_sheet).Cells(d, dough_name_read_col + 1) = dough_weight(d)
        Next d


''       ### 1.1.3 ### Check each dough for inside base dough
'
'        For d = 1 To (UBound(dough_name))
'
'            recipe_row = find_recipe_row(dough_name(d))
'            If recipe_row <> 404 Then
'                While Sheets(recipe_sheet).Cells(1, ingredient_read_col) <> "End of Ingredients" And ingredient_read_col < 500
'                    If Sheets(recipe_sheet).Cells(recipe_row, ingredient_read_col) <> "" Then
'                        Sheets(new_sheet_name).Cells(ingredient_fill_row, ingredient_name_fill_col) = Sheets(recipe_sheet).Cells(1, ingredient_read_col)
'                        Sheets(new_sheet_name).Cells(ingredient_fill_row, ingredient_percent_fill_col) = Sheets(recipe_sheet).Cells(recipe_row, ingredient_read_col)
'                        ingredient_total_percent = ingredient_total_percent + Sheets(new_sheet_name).Cells(ingredient_fill_row, ingredient_percent_fill_col)
'                        ingredient_fill_row = ingredient_fill_row + 1
'                    End If
'                    ingredient_read_col = ingredient_read_col + 1
'                Wend
'            End If
'        Next d



        
'       ### comment ### Create worksheets for mixing
        ingredient_fill_start_row = 13
        ingredient_name_fill_col = 2
        ingredient_percent_fill_col = 9
        ingredient_weight_fill_col = 5
        
        
        ReDim Preserve mixing_sheets(1 To 1)
        ReDim Preserve base_dough_name(1 To 1)
        ReDim Preserve base_dough_weight(1 To 1)
        
        insert_base_dough_offset = 0
        ReDim Preserve seek_terms(1 To 2)
        seek_terms(1) = "Dough, "
        seek_terms(2) = "Soaker, "
        
        s = 1
        While Sheets(temp_dough_sheet).Cells(s, dough_name_read_col) <> ""
            new_sheet_tag = "Mixing_"
            Dim tempname As String
            mixing_sheets(s - insert_base_dough_offset) = create_dough_sheet(dough_name(s - insert_base_dough_offset), dough_weight(s - insert_base_dough_offset), work_date, new_sheet_tag & (s - insert_base_dough_offset), ingredient_fill_start_row, ingredient_name_fill_col, ingredient_percent_fill_col, ingredient_weight_fill_col, recipe_sheet, template_name)
            base_dough = get_base_dough(mixing_sheets(s - insert_base_dough_offset), ingredient_fill_start_row, ingredient_name_fill_col, ingredient_weight_fill_col, seek_terms)
            If base_dough(1, 1) <> False Then
                For b = 1 To UBound(base_dough, 2) - 1
                    For r = 1 To 100
                        If Sheets(temp_dough_sheet).Cells(r, dough_name_read_col) = "" Then
                            Sheets(temp_dough_sheet).Rows(s).Insert shift:=xlUp, CopyOrigin:=xlFormatFromLeftOrAbove
                            Sheets(temp_dough_sheet).Cells(s, dough_name_read_col) = base_dough(1, b)
                            Sheets(temp_dough_sheet).Cells(s, dough_name_read_col + 1) = base_dough(2, b)
                            insert_base_dough_offset = insert_base_dough_offset + 1
                            s = s + 1
                            Exit For
                        ElseIf Sheets(temp_dough_sheet).Cells(r, dough_name_read_col) = base_dough(1, b) Then
                            Sheets(temp_dough_sheet).Cells(r, dough_name_read_col + 1) = Sheets(temp_dough_sheet).Cells(r, dough_name_read_col + 1) + base_dough(2, b)
                            Exit For
                        End If
                    Next r
                Next b

            End If
            
            ReDim Preserve mixing_sheets(1 To (UBound(mixing_sheets) + 1))
        s = s + 1
        Wend

        For s = 1 To 200
            If Sheets(temp_dough_sheet).Cells(s, dough_name_read_col) = "" Then
                Exit For
            Else
                ReDim Preserve dough_name(1 To s)
                ReDim Preserve dough_weight(1 To s)
                dough_name(s) = Sheets(temp_dough_sheet).Cells(s, dough_name_read_col)
                dough_weight(s) = Sheets(temp_dough_sheet).Cells(s, dough_name_read_col + 1)
            End If
        
        Next s
        
        ReDim Preserve mixing_sheets(1 To (UBound(mixing_sheets) - 1))
        Application.DisplayAlerts = False
        Sheets(mixing_sheets).Delete
        Sheets(temp_dough_sheet).Delete

        
        ReDim mixing_sheets(1 To 1)
        
        For s = 1 To UBound(dough_name)
            new_sheet_tag = "Mixing_"
            mixing_sheets(s) = create_dough_sheet(dough_name(s), dough_weight(s), work_date, new_sheet_tag & s, ingredient_fill_start_row, ingredient_name_fill_col, ingredient_percent_fill_col, ingredient_weight_fill_col, recipe_sheet, template_name)
            ReDim Preserve mixing_sheets(1 To (UBound(mixing_sheets) + 1))
        Next s
        
    ReDim Preserve mixing_sheets(1 To (UBound(mixing_sheets) - 1))
        
        For s = 1 To UBound(mixing_sheets)
            If InStr(Sheets(mixing_sheets(s)).Cells(4, 3), "Soaker") > 0 Then
                soaker_sheets(UBound(soaker_sheets)) = mixing_sheets(s)
                Sheets(mixing_sheets(s)).Cells(1, 2) = "Soaker"
                ReDim Preserve soaker_sheets(1 To (UBound(soaker_sheets) + 1))
            Else
                temp_sheets(UBound(temp_sheets)) = mixing_sheets(s)
                ReDim Preserve temp_sheets(1 To (UBound(temp_sheets) + 1))
            End If
        Next s
        
        ReDim Preserve temp_sheets(1 To (UBound(temp_sheets) - 1))
        ReDim mixing_sheets(1 To 1)
        mixing_sheets = compile_output_sheets(mixing_sheets, temp_sheets)
        ReDim Preserve mixing_sheets(1 To (UBound(mixing_sheets) - 1))
            
        If UBound(soaker_sheets) > 1 Then
            ReDim Preserve soaker_sheets(1 To (UBound(soaker_sheets) - 1))
        End If
        
        If do_mixing = True And do_starters = False Then
            Application.DisplayAlerts = False
            If soaker_sheets(1) <> "" Then
                Sheets(soaker_sheets).Delete
            End If
            
            Application.DisplayAlerts = True
        
        End If
        
        owners_sheet = "Owners_Doughs"
        
        For i = 1 To UBound(mixing_sheets)
            For o = 2 To 200
                If Sheets(mixing_sheets(i)).Cells(4, 3) = Sheets(owners_sheet).Cells(o, 1) Then
                    Sheets(mixing_sheets(i)).Cells(6, 3) = Sheets(owners_sheet).Cells(o, 2)
                    Sheets(mixing_sheets(i)).Cells(6, 7) = Sheets(owners_sheet).Cells(o, 3)
                    Exit For
                ElseIf Sheets(owners_sheet).Cells(o, 1) = "" Then
                    Exit For
                End If

            Next o
        
        Next i
        
        If do_mixing = True Then
            sheet_output = compile_output_sheets(sheet_output, mixing_sheets)
        End If
        

    End If
    
    
'   ### 2.0 ### calculate starters
    If do_starters = True Then
        Dim starters() As Variant, add_sheets As Integer
        Dim starter_sheets() As String
        ReDim starter_sheets(1 To 1)
        
        ReDim starters(1 To 2, 1 To 1)
        ReDim starter_name(1 To 1)
        ReDim starter_weight(1 To 1)
        ReDim seek_terms(1 To 1)
        
        starter_percent_feed = Sheets(control_sheet_name).Cells(8, 10)
        
        template_name = "Template_Starters"
        seek_terms(1) = "Starter, "
        For m = 1 To UBound(mixing_sheets)
        
            base_dough = get_base_dough(mixing_sheets(m), ingredient_fill_start_row, ingredient_name_fill_col, ingredient_weight_fill_col, seek_terms)
            If base_dough(1, 1) <> False Then
                For s = 1 To UBound(base_dough, 2) - 1
                    starter_name(UBound(starter_name)) = base_dough(1, s)
                    starter_weight(UBound(starter_weight)) = base_dough(2, s)
                    ReDim Preserve starter_name(1 To (UBound(starter_name) + 1))
                    ReDim Preserve starter_weight(1 To (UBound(starter_weight) + 1))
                Next s
            End If
        
        Next m
    
        For d = 1 To (UBound(starter_name) - 1)
            For s = 1 To (UBound(starters, 2))
                If starters(1, s) = "" Then
                    starters(1, s) = starter_name(d)
                    starters(2, s) = starter_weight(d)
                    ReDim Preserve starters(1 To 2, 1 To (UBound(starters, 2) + 1))
                ElseIf starters(1, s) = starter_name(d) Then
                    starters(2, s) = starters(2, s) + starter_weight(d)
                    Exit For
                End If
            Next s
        Next d
        
        Sheets.Add(Before:=Sheets(1)).name = temp_dough_sheet
        For d = 1 To (UBound(starters, 2) - 1)
            Sheets(temp_dough_sheet).Cells(d, dough_name_read_col) = starters(1, d)
            Sheets(temp_dough_sheet).Cells(d, dough_name_read_col + 1) = starters(2, d)
        Next d
    
        For d = 1 To 100
            If Sheets(temp_dough_sheet).Cells(d, dough_name_read_col) = "" Then
                Exit For
            ElseIf Sheets(temp_dough_sheet).Cells(d, dough_name_read_col) = "Starter, Levain" Then
                'starter_percent_feed = Mid(Sheets(temp_dough_sheet).Cells(d, dough_name_read_col), InStr(Sheets(temp_dough_sheet).Cells(d, dough_name_read_col), "%") - 2, 2) / 100
                Sheets(temp_dough_sheet).Cells(d, dough_name_read_col) = Sheets(temp_dough_sheet).Cells(d, dough_name_read_col) & ", " & starter_percent_feed * 100 & "%"
                starter_percent_total = 2 + starter_percent_feed
                
'                If InStr(Sheets(temp_dough_sheet).Cells(d, dough_name_read_col), "Starter, Levain, 60%") > 0 Then
'                    Sheets(temp_dough_sheet).Cells(d, dough_name_read_col + 1) = Sheets(temp_dough_sheet).Cells(d, dough_name_read_col + 1) + levain60_buffer
'                ElseIf InStr(Sheets(temp_dough_sheet).Cells(d, dough_name_read_col), "Starter, Levain, 20%") > 0 Then
'                    Sheets(temp_dough_sheet).Cells(d, dough_name_read_col + 1) = Sheets(temp_dough_sheet).Cells(d, dough_name_read_col + 1) + levain20_buffer
'                End If
                'Sheets(temp_dough_sheet).Cells(d, dough_name_read_col) = levain_shift & " " & Sheets(temp_dough_sheet).Cells(d, dough_name_read_col)
                Sheets(temp_dough_sheet).Cells(d, dough_name_read_col + 1) = Sheets(temp_dough_sheet).Cells(d, dough_name_read_col + 1) + levain60_buffer
                
                For v = 1 To levain_feeding_max - 1
                    Sheets(temp_dough_sheet).Rows(d).Insert shift:=xlUp, CopyOrigin:=xlFormatFromLeftOrAbove
                    Sheets(temp_dough_sheet).Cells(d, dough_name_read_col) = Sheets(temp_dough_sheet).Cells(d + v, dough_name_read_col) & " Feed #" & levain_feeding_max - v
                    Sheets(temp_dough_sheet).Cells(d, dough_name_read_col + 1) = Sheets(temp_dough_sheet).Cells(d + 1, dough_name_read_col + 1) / starter_percent_total * starter_percent_feed
                Next v
                d = d + v - 1
                Sheets(temp_dough_sheet).Cells(d, dough_name_read_col) = Sheets(temp_dough_sheet).Cells(d, dough_name_read_col) & " Feed #" & levain_feeding_max
            End If
        Next d
        
        ReDim starters(1 To 2, 1 To 1)
        
        For d = 1 To 100
            If Sheets(temp_dough_sheet).Cells(d, dough_name_read_col) = "" Then
                Exit For
            Else
                starters(1, d) = Sheets(temp_dough_sheet).Cells(d, dough_name_read_col)
                starters(2, d) = Sheets(temp_dough_sheet).Cells(d, dough_name_read_col + 1)
                ReDim Preserve starters(1 To 2, 1 To (UBound(starters, 2) + 1))
            End If
        Next d
    
        If starters(1, 1) <> "" Then
            ReDim Preserve starters(1 To 2, 1 To (UBound(starters, 2) - 1))
'        ReDim Preserve mixing_sheets(1 To (UBound(mixing_sheets) + 1))

        
'        add_sheets = UBound(mixing_sheets)
        
            For s = 1 To UBound(starters, 2)
                new_sheet_tag = "Starter_"
    '            mixing_sheets(add_sheets + s - 1) = create_dough_sheet(starters(1, s), starters(2, s), work_date, new_sheet_tag & s, ingredient_fill_start_row, ingredient_name_fill_col, ingredient_percent_fill_col, ingredient_weight_fill_col, recipe_sheet, template_name)
                starter_sheets(s) = create_dough_sheet(starters(1, s), starters(2, s), work_date, new_sheet_tag & s, ingredient_fill_start_row, ingredient_name_fill_col, ingredient_percent_fill_col, ingredient_weight_fill_col, recipe_sheet, template_name)
                
                If InStr(Sheets(starter_sheets(s)).Cells(4, dough_name_read_col), "Starter, Levain") > 0 Then
                    Sheets(starter_sheets(s)).Cells(4, dough_name_read_col) = levain_shift & " " & Sheets(starter_sheets(s)).Cells(4, dough_name_read_col)
                End If
                ReDim Preserve starter_sheets(1 To (UBound(starter_sheets) + 1))
    '            ReDim Preserve mixing_sheets(1 To (UBound(mixing_sheets) + 1))
            Next s

    '        ReDim Preserve mixing_sheets(1 To (UBound(mixing_sheets) - 1))
            ReDim Preserve starter_sheets(1 To (UBound(starter_sheets) - 1))
        

        
            owners_sheet = "Owners_Starters"
            For i = 1 To UBound(starter_sheets)
                For o = 2 To 200
                    If Sheets(starter_sheets(i)).Cells(4, 3) = Sheets(owners_sheet).Cells(o, 1) Then
                        Sheets(starter_sheets(i)).Cells(6, 3) = Sheets(owners_sheet).Cells(o, 2)
                        Sheets(starter_sheets(i)).Cells(6, 7) = Sheets(owners_sheet).Cells(o, 3)
                        Exit For
                    ElseIf Sheets(owners_sheet).Cells(o, 1) = "" Then
                        Exit For
                    End If
                Next o
            Next i
        
        End If
        
        Sheets(temp_dough_sheet).Delete
        
        If soaker_sheets(1) <> "" Then
            For i = 1 To UBound(soaker_sheets)
                For o = 2 To 200
                    If Sheets(soaker_sheets(i)).Cells(4, 3) = Sheets(owners_sheet).Cells(o, 1) Then
                        Sheets(soaker_sheets(i)).Cells(6, 3) = Sheets(owners_sheet).Cells(o, 2)
                        Sheets(soaker_sheets(i)).Cells(6, 7) = Sheets(owners_sheet).Cells(o, 3)
                        Exit For
                    ElseIf Sheets(owners_sheet).Cells(o, 1) = "" Then
                        Exit For
                    End If
                Next o
            Next i
        End If
        
        If do_starters = True Then
            If starters(1, 1) <> "" Then
                sheet_output = compile_output_sheets(sheet_output, starter_sheets)
            End If
            If soaker_sheets(1) <> "" Then
                sheet_output = compile_output_sheets(sheet_output, soaker_sheets)
            End If
        End If
        
        If do_mixing = False Then
            Application.DisplayAlerts = False
            Sheets(mixing_sheets).Delete
            Application.DisplayAlerts = True
        End If
        

    End If
    
    
'   ### 3.0 ### Generate Processing work orders


    If do_processing = True Then
            
    
    bread_tag = "FI, "
    
        
    search_max = 50
    no_bread_count = 0
    bread_name_read_row = 11
    bread_name_read_col = 3
    processing_template_name = "Template_Process_Bake"
    dough_tag = "Dough, "
    bread_count = 0
    ReDim Preserve bread_sheets(1 To 1)
    
        
        While no_bread_count < search_max
            If InStr(Sheets(source_sheet).Cells(bread_name_read_row, bread_name_read_col), dough_tag) > 0 Then
                bread_dough = Sheets(source_sheet).Cells(bread_name_read_row, bread_name_read_col)
            End If
            
            
            If InStr(Sheets(source_sheet).Cells(bread_name_read_row, bread_name_read_col), bread_tag) > 0 Then
                bread_count = bread_count + 1
                new_sheet_name = "Processing " & bread_count
                bread_sheets(UBound(bread_sheets)) = new_sheet_name
                ReDim Preserve bread_sheets(1 To UBound(bread_sheets) + 1)
                
                Sheets(processing_template_name).Copy Before:=Sheets(1)
                Sheets(processing_template_name & " (2)").name = new_sheet_name
                Sheets(new_sheet_name).Cells(5, 3) = bread_dough
                'Sheets(new_sheet_name).Cells(3, 3) = work_date
                Sheets(new_sheet_name).Cells(4, 3) = Right(Sheets(source_sheet).Cells(bread_name_read_row, bread_name_read_col), Len(Sheets(source_sheet).Cells(bread_name_read_row, bread_name_read_col)) - 5)
                Sheets(new_sheet_name).Cells(6, 3) = Sheets(source_sheet).Cells(bread_name_read_row, 16) * 1000
                ' ### comment ### comment the following line out for blank quantity
                'Sheets(new_sheet_name).Cells(7, 3) = Sheets(source_sheet).Cells(bread_name_read_row, 9)
                
                no_bread_count = 0
            End If
            
            If Sheets(source_sheet).Cells(bread_name_read_row, bread_name_read_col) = "" Then
                no_bread_count = no_bread_count + 1
            End If
            
            bread_name_read_row = bread_name_read_row + 1

        Wend
        
        ReDim Preserve bread_sheets(1 To UBound(bread_sheets) - 1)
        sheet_output = compile_output_sheets(sheet_output, bread_sheets)
    
    End If
    ' ### end 3.0 ###
    
'   ### 4.0 ### Generate Baking work orders


    If do_baking = True Then
            

    
    bread_tag = "FI, "
    
        
    search_max = 50
    no_bread_count = 0
    bread_name_read_row = 14
    bread_name_read_col = 3
    processing_template_name = "Template_Baking"
    dough_tag = "Dough, "
    bread_count = 0
    ReDim Preserve baking_sheets(1 To 1)
    
        
        While no_bread_count < search_max
            If InStr(Sheets(source_sheet).Cells(bread_name_read_row, bread_name_read_col), dough_tag) > 0 Then
                bread_dough = Sheets(source_sheet).Cells(bread_name_read_row, bread_name_read_col)
            End If
            
            
            If InStr(Sheets(source_sheet).Cells(bread_name_read_row, bread_name_read_col), bread_tag) > 0 Then
                bread_count = bread_count + 1
                new_sheet_name = "Baking " & bread_count
                baking_sheets(UBound(baking_sheets)) = new_sheet_name
                ReDim Preserve baking_sheets(1 To UBound(baking_sheets) + 1)
                
                Sheets(processing_template_name).Copy Before:=Sheets(1)
                Sheets(processing_template_name & " (2)").name = new_sheet_name
                'Sheets(new_sheet_name).Cells(4, 3) = bread_dough
                Sheets(new_sheet_name).Cells(3, 3) = work_date
                Sheets(new_sheet_name).Cells(4, 3) = Right(Sheets(source_sheet).Cells(bread_name_read_row, bread_name_read_col), Len(Sheets(source_sheet).Cells(bread_name_read_row, bread_name_read_col)) - 5)
                Sheets(new_sheet_name).Cells(5, 3) = Sheets(source_sheet).Cells(bread_name_read_row, 16) * 1000
                ' ### comment ### comment the following line out for blank quantity
                Sheets(new_sheet_name).Cells(6, 3) = Sheets(source_sheet).Cells(bread_name_read_row, 9)
                
                no_bread_count = 0
            End If
            
            If Sheets(source_sheet).Cells(bread_name_read_row, bread_name_read_col) = "" Then
                no_bread_count = no_bread_count + 1
            End If
            
            bread_name_read_row = bread_name_read_row + 1

        Wend
        
        ReDim Preserve baking_sheets(1 To UBound(baking_sheets) - 1)
        sheet_output = compile_output_sheets(sheet_output, baking_sheets)
    
    End If
    ' ### end 4.0 ###
    

'    If do_mixing = True Then
'        Sheets(mixing_sheet).Move
'    End If
    
'    ReDim Preserve sheet_output(1 To (UBound(sheet_output) - 1))
    
    If sheet_output(1) <> "" Then
    
        ReDim Preserve sheet_output(1 To (UBound(sheet_output) - 1))
        For o = 1 To UBound(sheet_output)
            Sheets(sheet_output(o)).Visible = True
        Next o
    
        Sheets(sheet_output).Move
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Function create_dough_sheet(dough_name As Variant, dough_weight As Variant, work_date As String, new_sheet_name As String, ingredient_fill_start_row As Integer, ingredient_name_fill_col As Integer, ingredient_percent_fill_col As Integer, ingredient_weight_fill_col As Integer, recipe_sheet As String, template As String) As String

    Dim recipe_row, ingredient_read_col, ingredient_fill_row, ingredient_total_rows As Integer
    Dim ingredient_total_percent As Double
    

    Sheets(template).Copy Before:=Sheets(1)
    Sheets(template & " (2)").name = new_sheet_name
    Sheets(new_sheet_name).Cells(4, 3) = dough_name
    Sheets(new_sheet_name).Cells(3, 3) = work_date
    
    If InStr(dough_name, "Feed #") > 0 Then
        dough_name = Left(dough_name, InStr(dough_name, "Feed #") - 2)
    End If
    
    recipe_row = find_recipe_row(dough_name, recipe_sheet)
    ingredient_total_rows = 17
    ingredient_read_col = 2
    ingredient_fill_row = ingredient_fill_start_row

    If recipe_row <> 404 Then
        While Sheets(recipe_sheet).Cells(1, ingredient_read_col) <> "End of Ingredients" And ingredient_read_col < 500
            If Sheets(recipe_sheet).Cells(recipe_row, ingredient_read_col) <> "" Then
                Sheets(new_sheet_name).Cells(ingredient_fill_row, ingredient_name_fill_col) = Sheets(recipe_sheet).Cells(1, ingredient_read_col)
                Sheets(new_sheet_name).Cells(ingredient_fill_row, ingredient_percent_fill_col) = Sheets(recipe_sheet).Cells(recipe_row, ingredient_read_col)

                ingredient_total_percent = ingredient_total_percent + Sheets(new_sheet_name).Cells(ingredient_fill_row, ingredient_percent_fill_col)
                ingredient_fill_row = ingredient_fill_row + 1
            End If
            ingredient_read_col = ingredient_read_col + 1
        Wend
    End If
    
    For r = ingredient_fill_start_row To ingredient_fill_start_row + ingredient_total_rows
        If ingredient_total_percent <> 0 And Sheets(new_sheet_name).Cells(r, ingredient_percent_fill_col) <> "" Then
            Sheets(new_sheet_name).Cells(r, ingredient_weight_fill_col) = dough_weight / ingredient_total_percent * Sheets(new_sheet_name).Cells(r, ingredient_percent_fill_col)
        ElseIf ingredient_total_percent = 0 Then
            Sheets(new_sheet_name).Cells(r, ingredient_weight_fill_col) = dough_weight
            Exit For
        Else
            Exit For
        End If
    Next r
    
'    Application.Wait Now + TimeValue("00:00:02")
    create_dough_sheet = new_sheet_name

End Function

Function find_recipe_row(dough_name As Variant, recipe_sheet As String) As Integer

    Dim recipe_row, recipe_col As Integer
    recipe_row = 1
    recipe_col = 1

    For r = 1 To 200
        If Sheets(recipe_sheet).Cells(r, recipe_col) = dough_name Then
            recipe_row = r
            Exit For
        ElseIf Sheets(recipe_sheet).Cells(r, recipe_col) = "End of Doughs" Then
            recipe_row = 404
            Exit For
        End If 'a
    Next r

    find_recipe_row = recipe_row

End Function


Function get_base_dough(new_sheet_name As String, ingredient_fill_start_row As Integer, ingredient_name_fill_col As Integer, ingredient_weight_fill_col As Integer, seek_terms As Variant) As Variant

    Dim ingredient_total_rows As Integer
    Dim base_dough() As Variant
    ingredient_total_rows = 17
    ReDim base_dough(1 To 2, 1 To 1)

    For t = 1 To UBound(seek_terms)
        For r = 1 To ingredient_total_rows
            If InStr(Sheets(new_sheet_name).Cells(r - 1 + ingredient_fill_start_row, ingredient_name_fill_col), seek_terms(t)) > 0 Then
                base_dough(1, UBound(base_dough, 2)) = Left(Sheets(new_sheet_name).Cells(r - 1 + ingredient_fill_start_row, ingredient_name_fill_col), InStr(Sheets(new_sheet_name).Cells(r - 1 + ingredient_fill_start_row, ingredient_name_fill_col), "[") - 2)
                base_dough(2, UBound(base_dough, 2)) = Sheets(new_sheet_name).Cells(r - 1 + ingredient_fill_start_row, ingredient_weight_fill_col)
                ReDim Preserve base_dough(1 To 2, 1 To (UBound(base_dough, 2) + 1))
 '               base_dough(1, 1) = True
                
    '            base_dough(0) = True
    '            Exit For
            Else
'                base_dough(1, 1) = False
            End If
        Next r
    Next t
    
'    ReDim Preserve base_dough(1 To 2, 1 To (UBound(base_dough, 2) - 1))
    
    If base_dough(1, 1) = "" Then
        base_dough(1, 1) = False
    End If
    
    get_base_dough = base_dough
    
End Function

Function combine_duplicates(name_array)

    


End Function

Function compile_output_sheets(sheet_output As Variant, sheet_input As Variant) As Variant
    
    For i = 1 To UBound(sheet_input)
        sheet_output(UBound(sheet_output)) = sheet_input(i)
        ReDim Preserve sheet_output(1 To UBound(sheet_output) + 1)
    Next i


    compile_output_sheets = sheet_output

End Function

