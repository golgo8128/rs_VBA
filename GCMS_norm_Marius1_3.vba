Option Explicit

Const BLOCK_ColNum As Integer = 2
Const SampleName_ColNum As Integer = 4


Sub Marius_calc_qc_stats1_3()

    Dim qc_block_mean() As Double
    Dim qc_mean As Double
    Dim data_sheet As Worksheet
    Dim qc_mean_sheet As Worksheet
    Dim normalized_sheet As Worksheet
    Dim i As Long
    Dim j As Long
    Dim max_row As Long
    Dim max_col As Long
    Dim max_block_num As Long
    Dim cur_block_num As Long
    
    Set data_sheet = ActiveSheet
    max_row = data_sheet.Range("A1").End(xlDown).Row
    max_col = data_sheet.Range("A1").End(xlToRight).Column
    
    max_block_num = Application.WorksheetFunction.Max(Range(ActiveSheet.Cells(2, 2), _
                                                      Cells(ActiveSheet.Range("A1").End(xlDown).Row, BLOCK_ColNum)))
    
    Worksheets.Add after:=data_sheet
    Set qc_mean_sheet = ActiveSheet
    qc_mean_sheet.Name = find_unused_sheetname(data_sheet.Name & "_QC_means_")
            
    Worksheets.Add after:=qc_mean_sheet
    Set normalized_sheet = ActiveSheet
    normalized_sheet.Name = find_unused_sheetname(data_sheet.Name & "_norm_")
            
    data_sheet.Activate
            
    qc_mean_sheet.Cells(2, 1) = "All blocks"
    
    For i = 1 To max_block_num
         qc_mean_sheet.Cells(i + 2, 1) = i
    Next i
    
    For i = 1 To max_row
        For j = 1 To SampleName_ColNum
            normalized_sheet.Cells(i, j) = data_sheet.Cells(i, j)
        Next j
    Next i
    
    For j = SampleName_ColNum + 1 To max_col
    
        qc_block_mean = marius_quality_control_mean_block(data_sheet, j)
        qc_mean = marius_quality_control_mean(data_sheet, j)
            
        qc_mean_sheet.Cells(1, j - SampleName_ColNum + 1) = data_sheet.Cells(1, j)
        qc_mean_sheet.Cells(2, j - SampleName_ColNum + 1) = qc_mean
    
        For i = 1 To UBound(qc_block_mean) ' max_block_num
            qc_mean_sheet.Cells(i + 2, j - SampleName_ColNum + 1) = qc_block_mean(i)
        Next i
        
        cur_block_num = 1
        normalized_sheet.Cells(1, j) = data_sheet.Cells(1, j)
        For i = 2 To max_row
            If data_sheet.Cells(i, BLOCK_ColNum) <> "" Then
                cur_block_num = data_sheet.Cells(i, BLOCK_ColNum).Value
            End If
            
            normalized_sheet.Cells(i, j) = data_sheet.Cells(i, j) * qc_mean / qc_block_mean(cur_block_num)
        
        Next i
    
    Next j

End Sub


Function marius_quality_control_mean(ByRef data_sheet As Worksheet, ByVal icolnum As Integer) As Double

  Dim block_total As Double
  Dim cell_count As Long
  Dim i As Long
    
  block_total = 0
  cell_count = 0
  
  For i = 2 To data_sheet.Range("A1").End(xlDown).Row
          
    If marius_quality_control_str_judge(data_sheet.Cells(i, SampleName_ColNum)) = True Then
        block_total = block_total + data_sheet.Cells(i, icolnum).Value
        cell_count = cell_count + 1
        With data_sheet.Cells(i, SampleName_ColNum).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
    End With
        
        
    End If
    
  Next i
  
  marius_quality_control_mean = block_total / cell_count
  

End Function




Function marius_quality_control_mean_block(ByRef data_sheet As Worksheet, ByVal icolnum As Integer) As Double()

  Dim max_block_num As Integer
  Dim cur_block_num As Integer ' Block number should start with 1
  Dim pre_block_num As Integer
  Dim block_total As Double
  Dim cell_count As Long
  Dim i As Long
    
  Dim ret_arr() As Double
    
  max_block_num = Application.WorksheetFunction.Max(Range(data_sheet.Cells(BLOCK_ColNum, 2), _
                                                          data_sheet.Cells(data_sheet.Range("A1").End(xlDown).Row, BLOCK_ColNum)))
  
  ReDim ret_arr(max_block_num)

  pre_block_num = 1
  block_total = 0
  cell_count = 0
  
  For i = 2 To data_sheet.Range("A1").End(xlDown).Row
    If data_sheet.Cells(i, BLOCK_ColNum).Value = "" Then
        cur_block_num = pre_block_num
    Else
        cur_block_num = data_sheet.Cells(i, BLOCK_ColNum).Value
    End If
    
    If pre_block_num <> cur_block_num Then
        ret_arr(pre_block_num) = block_total / cell_count
        block_total = 0
        cell_count = 0
    End If
          
    If marius_quality_control_str_judge(data_sheet.Cells(i, SampleName_ColNum)) = True Then
        block_total = block_total + data_sheet.Cells(i, icolnum).Value
        cell_count = cell_count + 1
    End If
    
    pre_block_num = cur_block_num
    
  Next i
  
  ret_arr(pre_block_num) = block_total / cell_count
  
  marius_quality_control_mean_block = ret_arr
  

End Function


Function marius_quality_control_str_judge(ByVal istr As String) As Boolean

    Dim i As Long

    marius_quality_control_str_judge = True

    If InStr(istr, "QC") Then
        For i = InStr(istr, "QC") + 2 To Len(istr)
            If InStr("0123456789 ", Mid(istr, i, 1)) = 0 Then
                marius_quality_control_str_judge = False
                Exit For
            End If
        Next i
    Else
       marius_quality_control_str_judge = False
    End If


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


