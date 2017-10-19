Option Explicit
Option Base 1

Const Sample_Abbr_ColNum As Integer = 1


Sub GCMS_datcollect_Manjula1_2()
 
    Dim i, j, k As Integer
    Dim metab_count, same_metab_count As Integer
    Dim data_sheet, dat_collect_sheet As Worksheet
    Dim calib_amt_row_start As Integer
    Dim peak_name, prev_peak_name As String

    Dim Peak_Name_ColNum As Integer ' = 4
    Dim Amt_ColNum As Integer ' = 7
    Dim Calib_Amt_ColNum As Integer ' = 9

    Dim match_colnams(3) As String
    Dim match_cols_row() As Long
    Dim match_row As Long

    match_colnams(1) = "Peak Name"
    match_colnams(2) = "Amt"
    match_colnams(3) = "Calib Amt"

    match_cols_row = find_first_row_that_contains_words1(match_colnams, 1, 100, 1, 25)
    match_row = match_cols_row(UBound(match_colnams) + 1)

    If match_row = 0 Then
        MsgBox "Cannot find a row with all of the keywords: " & Join(match_colnams, ", "), vbCritical
        Exit Sub
    End If

    For k = 1 To UBound(match_colnams)
        
        With Cells(match_row, match_cols_row(k)).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.399975585192419
            .PatternTintAndShade = 0
        End With
        
    Next k
    
    Peak_Name_ColNum = match_cols_row(1)
    Amt_ColNum = match_cols_row(2)
    Calib_Amt_ColNum = match_cols_row(3)

    Set data_sheet = ActiveSheet

    If Len(data_sheet.Name) > 20 Then
        MsgBox "Length of the sheet name should not go over 20.", vbCritical
        Exit Sub
    End If

    Worksheets.Add after:=data_sheet
    Set dat_collect_sheet = ActiveSheet
    dat_collect_sheet.Name = find_unused_sheetname(data_sheet.Name & "_collect_")
            

    data_sheet.Activate


    For i = 1 To 500
        If Cells(i, Calib_Amt_ColNum) = "Calib Amt" Then
            Cells(i, Calib_Amt_ColNum).Select
            calib_amt_row_start = i + 1
            Exit For
        End If
    Next i

    prev_peak_name = ""
    metab_count = 0

    i = calib_amt_row_start
    Do While Not IsEmpty(Cells(i, Calib_Amt_ColNum))
        If Cells(i, Calib_Amt_ColNum) = "---" Then
            With Cells(i, Calib_Amt_ColNum).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent4
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With

        
            peak_name = Cells(i, Peak_Name_ColNum)
            
            If peak_name <> prev_peak_name Then
                same_metab_count = 1
                metab_count = metab_count + 1
                dat_collect_sheet.Cells(metab_count + 1, 1) = peak_name
            Else
                same_metab_count = same_metab_count + 1
            End If
        
            If metab_count = 1 Then
                dat_collect_sheet.Cells(1, same_metab_count + 1) = Cells(i, Sample_Abbr_ColNum)
            End If
        
            dat_collect_sheet.Cells(metab_count + 1, same_metab_count + 1) = Cells(i, Amt_ColNum)
        
            prev_peak_name = peak_name
        
        End If
        
        i = i + 1
    Loop
    
    
    dat_collect_sheet.Activate
    Columns("A:A").EntireColumn.AutoFit


End Sub



Function find_first_row_that_contains_words1( _
    ByRef _
    match_words() As String, _
    ByRef _
    row_from As Long, _
    row_to As Long, _
    col_from As Long, _
    col_to As Long) As Long()

    Dim matched_col As Long
    Dim matched_cols_row() As Long
    Dim matched_row As Long
    Dim num_matched_words As Integer
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    ReDim matched_cols_row(UBound(match_words) + 1) ' The last item is matched row

    matched_row = 0

    For i = row_from To row_to
        num_matched_words = 0
        For k = 1 To UBound(match_words)
            matched_col = 0
            For j = col_from To col_to
                If Cells(i, j) = match_words(k) Then
                    matched_col = j
                    Exit For
                End If
            Next j
            If matched_col > 0 Then
                matched_cols_row(k) = matched_col
                num_matched_words = num_matched_words + 1
            Else
                Exit For
            End If
        Next k
        
        If num_matched_words = UBound(match_words) Then
            matched_row = i
            Exit For
        End If
    Next i

    matched_cols_row(UBound(match_words) + 1) = matched_row
    find_first_row_that_contains_words1 = matched_cols_row


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


