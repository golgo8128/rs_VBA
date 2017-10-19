Option Explicit
Option Base 1

Const IsomDeri_info_sheetname As String = "IsomersDerivatives_info"


Sub IntegIsomerDerivs_Manjula1_1()
'
'

    Dim data_sheet As Worksheet
    Dim rowcombi_sheet As Worksheet
    Dim rowcombi_sheet_row As Long

    Dim tmp_syno_names(2) As String
    Dim tmp_rows() As Long
    Dim tmpi As Integer

    Set data_sheet = ActiveSheet
    
    Worksheets.Add after:=data_sheet
    Set rowcombi_sheet = ActiveSheet
    rowcombi_sheet.Name = find_unused_sheetname(data_sheet.Name & "_Combined_")
    
    data_sheet.Activate
    
    If Cells(1, 2) <> "Converted name" Then
        Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Cells(1, 2) = "Converted name"
    End If

    tmp_syno_names(1) = "Aconitic Acid Tri TMS 01"
    tmp_syno_names(2) = "Aconitic Acid Tri TMS 02"

    If Not sheetname_exists(IsomDeri_info_sheetname) Then
        
        MsgBox "Data sheet """ & IsomDeri_info_sheetname & _
               """ not found in the current workbook", vbExclamation
        Exit Sub
        
    End If

    tmp_rows = search_syno_metab_rows(data_sheet, _
                                      tmp_syno_names, "Aconitic acid")
    tmpi = tmp_rows(1)
    tmpi = tmp_rows(2)
    tmpi = UBound(tmp_rows)

    rowcombi_sheet_row = 3

    syno_rows_sum_to_other_sheet data_sheet, tmp_rows, rowcombi_sheet, rowcombi_sheet_row
    rowcombi_sheet.Cells(rowcombi_sheet_row, 2) = "Aconitic acid"
    
End Sub

Sub syno_rows_sum_to_other_sheet(ByRef data_sheet As Worksheet, _
                                 ByRef data_sheet_rows() As Long, _
                                 ByRef rowcombi_sheet As Worksheet, _
                                 ByVal rowcombi_sheet_row As Long)
    Dim i As Long
    Dim j As Long
    Dim colsum As Long
    Dim max_col As Long
    
    max_col = data_sheet.Range("A1").End(xlToRight).Column

    For j = 3 To max_col
    
        colsum = 0
        For i = 1 To UBound(data_sheet_rows)
            colsum = colsum + data_sheet.Cells(data_sheet_rows(i), j)
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

    max_row = data_sheet.Range("A1").End(xlDown).Row


    syno_metab_rows_ct = 0
    For i = 1 To max_row ' Should start from 2?
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

    ReDim osyno_metab_rows(syno_metab_rows_ct)

    For k = 1 To syno_metab_rows_ct
        osyno_metab_rows(k) = syno_metab_rows(k)
    Next k
    
    search_syno_metab_rows = osyno_metab_rows
    

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
