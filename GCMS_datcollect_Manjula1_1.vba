Option Explicit
Option Base 1

Const Sample_Abbr_ColNum As Integer = 1
Const Peak_Name_ColNum As Integer = 4
Const Amt_ColNum As Integer = 7
Const Calib_Amt_ColNum As Integer = 9

Sub GCMS_datcollect_Manjula1_1()
 
    Dim i, j As Integer
    Dim metab_count, same_metab_count As Integer
    Dim data_sheet, dat_collect_sheet As Worksheet
    Dim calib_amt_row_start As Integer
    Dim peak_name, prev_peak_name As String

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

