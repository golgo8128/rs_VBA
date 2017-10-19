Option Explicit
Option Base 1

Const IsomDeri_info_sheetname As String = "IsomersDerivatives_info"


Sub IntegIsomerDerivs_Manjula1_2()
'
'

    Dim data_sheet As Worksheet
    Dim rowcombi_sheet As Worksheet
    Dim rowcombi_sheet_row As Long
    Dim IsomDeri_info_sheet As Worksheet
    Dim IsomDeri_info_sheet_row As Long
    
    Dim syno_found_rows() As Long
    Dim cname As String
    Dim synos() As String

    Dim i As Long
    Dim j As Long

    Set data_sheet = ActiveSheet
    
    If Not sheetname_exists(IsomDeri_info_sheetname) Then
        
        MsgBox "Sheet """ & IsomDeri_info_sheetname & _
               """ not found in the current workbook", vbExclamation
        Exit Sub
    Else
        Set IsomDeri_info_sheet = get_sheet_from_name(IsomDeri_info_sheetname)
    End If
    
    If IsEmpty(data_sheet.Cells(2, 1)) Or IsEmpty(data_sheet.Cells(2, 1)) Then
        MsgBox "Invalid data", vbExclamation
        Exit Sub
    End If
    
    Worksheets.Add after:=data_sheet
    Set rowcombi_sheet = ActiveSheet
    rowcombi_sheet.Name = find_unused_sheetname(Left(data_sheet.Name, 23) & "_Combi")
    rowcombi_sheet.Cells(1, 1) = "Original name(s)"
    rowcombi_sheet.Cells(1, 2) = "Converted name"
    
    data_sheet.Activate
    
    If data_sheet.Cells(1, 2) <> "Converted name" Then
        data_sheet.Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        data_sheet.Cells(1, 2) = "Converted name"
    End If

    For j = 3 To data_sheet.Range("B1").End(xlToRight).Column
        rowcombi_sheet.Cells(1, j) = data_sheet.Cells(1, j)
    Next j

    rowcombi_sheet_row = 2
    IsomDeri_info_sheet_row = 2
    Do While Not IsEmpty(IsomDeri_info_sheet.Cells(IsomDeri_info_sheet_row, 1))
        cname = IsomDeri_info_sheet.Cells(IsomDeri_info_sheet_row, 1)
        synos = get_row_items(IsomDeri_info_sheet, IsomDeri_info_sheet_row)
        syno_found_rows = search_syno_metab_rows(data_sheet, _
                                                 synos, cname)
        If syno_found_rows(1) > 0 Then
            syno_rows_sum_to_other_sheet data_sheet, syno_found_rows, rowcombi_sheet, rowcombi_sheet_row
            rowcombi_sheet.Cells(rowcombi_sheet_row, 2) = cname
            rowcombi_sheet_row = rowcombi_sheet_row + 1
        End If
        
        IsomDeri_info_sheet_row = IsomDeri_info_sheet_row + 1
    Loop


End Sub


Function get_row_items(ByRef IsomDeri_info_sheet As Worksheet, irow As Long) As String()

    Dim j As Integer
    Dim tmpstrarray(256) As String
    Dim ostrarray() As String
    Dim ct As Long
    
    ct = 0
    j = 3
    Do While Not IsEmpty(IsomDeri_info_sheet.Cells(irow, j))
        ct = ct + 1
        tmpstrarray(ct) = IsomDeri_info_sheet.Cells(irow, j)
        j = j + 1
    Loop
    
    ReDim ostrarray(ct)
    
    For j = 1 To ct
        ostrarray(j) = tmpstrarray(j)
    Next j

    get_row_items = ostrarray
    

End Function



Sub syno_rows_sum_to_other_sheet(ByRef data_sheet As Worksheet, _
                                 ByRef data_sheet_rows() As Long, _
                                 ByRef rowcombi_sheet As Worksheet, _
                                 ByVal rowcombi_sheet_row As Long)
    Dim i As Long
    Dim j As Long
    Dim colsum As Double
    Dim max_col As Long
    
    max_col = data_sheet.Range("B1").End(xlToRight).Column

    For j = 3 To max_col
    
        colsum = 0
        For i = 1 To UBound(data_sheet_rows)
        
            If j = 3 Then
                If Not IsEmpty(rowcombi_sheet.Cells(rowcombi_sheet_row, 1)) Then
                    rowcombi_sheet.Cells(rowcombi_sheet_row, 1) = _
                        rowcombi_sheet.Cells(rowcombi_sheet_row, 1) & "; "
                End If
                rowcombi_sheet.Cells(rowcombi_sheet_row, 1) = _
                    rowcombi_sheet.Cells(rowcombi_sheet_row, 1) & data_sheet.Cells(data_sheet_rows(i), 1)
            End If
            On Error GoTo SUMFAIL
            colsum = colsum + data_sheet.Cells(data_sheet_rows(i), j)
            If IsEmpty(data_sheet.Cells(data_sheet_rows(i), j)) Then
                With data_sheet.Cells(data_sheet_rows(i), j).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If

            GoTo SUMOK
            
SUMFAIL:
            With data_sheet.Cells(data_sheet_rows(i), j).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Resume Next
SUMOK:
        Next i
        
        rowcombi_sheet.Cells(rowcombi_sheet_row, j) = colsum
        
    Next j



End Sub

Function search_syno_metab_rows(ByRef data_sheet As Worksheet, _
                                ByRef syno_names() As String, _
                                ByVal conv_name As String) As Long()

    Dim i As Long
    Dim k As Long
    Dim max_row As Long
    Dim syno_metab_rows_ct As Integer
    Dim syno_metab_rows(256), osyno_metab_rows() As Long

    max_row = data_sheet.Range("A2").End(xlDown).Row


    syno_metab_rows_ct = 0
    For i = 2 To max_row ' Should start from 2?
        For k = 1 To UBound(syno_names)
            If data_sheet.Cells(i, 1) = syno_names(k) Then ' First column should be metabolite names
                syno_metab_rows_ct = syno_metab_rows_ct + 1
                syno_metab_rows(syno_metab_rows_ct) = i
                
                With data_sheet.Cells(i, 1).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 10092543
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                
                data_sheet.Cells(i, 2) = conv_name
                
                Exit For
            End If
        Next k
    Next i

    If syno_metab_rows_ct > 0 Then
        ReDim osyno_metab_rows(syno_metab_rows_ct)
    
        For k = 1 To syno_metab_rows_ct
            osyno_metab_rows(k) = syno_metab_rows(k)
        Next k
    Else
        ReDim osyno_metab_rows(1)
        osyno_metab_rows(1) = 0
    End If
    
    search_syno_metab_rows = osyno_metab_rows
    

End Function


Function get_sheet_from_name(ByVal isheetname As String) As Worksheet


    Dim i As Long
    Dim flag As Boolean
    
    flag = False
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = isheetname Then
            flag = True
            Set get_sheet_from_name = Worksheets(i)
            Exit For
        End If
    Next i

End Function

Function sheetname_exists(ByVal isheetname As String) As Boolean


    Dim i As Long
    Dim flag As Boolean
    
    flag = False
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = isheetname Then
            flag = True
            Exit For
        End If
    Next i

    sheetname_exists = flag

End Function

Function find_unused_sheetname(ByVal isheetname_head As String) As String

    Dim i As Long
    Dim new_sheetname_cand As String
    Dim new_sheetname As String
    
    new_sheetname = ""
    
    For i = 1 To 100
        new_sheetname_cand = isheetname_head & CStr(i)
        If Not sheetname_exists(new_sheetname_cand) Then
            new_sheetname = new_sheetname_cand
            Exit For
        End If
        
    Next i

    If new_sheetname = "" Then
        Err.Raise 601, "User error", "Couldn't generate new sheet name"
    End If

    find_unused_sheetname = new_sheetname

End Function
