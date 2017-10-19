Option Explicit
Option Base 1

Sub SelectedCols_strjoin1_1()
'
' Macro1 Macro
'

'
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim selected_cols_redu() As Integer
    Dim selected_cols() As Integer
    Dim testarray(6) As Integer
    Dim SelectedArea As Range
    Dim rCell As Range
    Set SelectedArea = Selection
    Dim write_col As Integer
    Dim ostr As String
    Dim ostr_add As String
    
    ReDim selected_cols_redu(SelectedArea.Count)
    
    i = 1
    For Each rCell In SelectedArea
        selected_cols_redu(i) = rCell.Column
        i = i + 1
    Next rCell

    Dim otestarray() As Integer

    selected_cols = nonredu_int_array(selected_cols_redu)

    For j = 1 To UBound(selected_cols)
        With Cells(1, selected_cols(j)).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    Next j

    j = 1
    Do While Not IsEmpty(Cells(1, j))
        j = j + 1
    Loop

    write_col = j

    i = 2
    Do While Not IsEmpty(Cells(i, 1))
        ostr = ""
        For k = 1 To UBound(selected_cols)
            j = selected_cols(k)
            If Cells(i, j) <> "" Then
                ostr_add = Cells(1, j) & ": " & Cells(i, j)
                If ostr = "" Then
                    ostr = ostr_add
                Else
                    ostr = ostr & "; " & ostr_add
                End If
            End If
        Next k
        
        If Not IsEmpty(ostr) Then
            Cells(i, write_col) = ostr
        End If
        
        With Cells(i, write_col).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With

        With Cells(i, 1).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With

        i = i + 1
    Loop


End Sub


Function nonredu_int_array(ByRef iarray() As Integer) As Integer()
    ' slow ...

    Dim i As Integer
    Dim j As Integer
    Dim input_size As Integer
    Dim nonredu_num As Integer
    Dim marray() As Integer
    Dim oarray() As Integer
    Dim already_exist_flag As Boolean

    ReDim marray(UBound(iarray))

    nonredu_num = 0
    For i = 1 To UBound(iarray)
        
        already_exist_flag = False
        
        For j = 1 To nonredu_num
            If iarray(i) = marray(j) Then
                already_exist_flag = True
                Exit For
            End If
        Next j
            
        If already_exist_flag = False Then
            nonredu_num = nonredu_num + 1
            marray(nonredu_num) = iarray(i)
        End If
        
    Next i

    ReDim oarray(nonredu_num)

    For i = 1 To nonredu_num
        oarray(i) = marray(i)
    Next i
    
    nonredu_int_array = oarray

End Function
