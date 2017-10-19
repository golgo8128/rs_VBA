Attribute VB_Name = "Module1"
Option Explicit

Sub HandyWord_Sort1()
Attribute HandyWord_Sort1.VB_ProcData.VB_Invoke_Func = " _n14"
'
' Macro2 Macro
'

'
    Dim row_c As Integer

    Columns("A:C").Select
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("B2:B10000"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("C2:C10000"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("A2:A10000"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A1:C10000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    row_c = 2
    
    Do While Cells(row_c, 1) <> ""
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(row_c, 1), Address:= _
        "http://ejje.weblio.jp/content/" & Cells(row_c, 1)
        If Cells(row_c, 3) = "" Then
            Cells(row_c, 3) = Now
        End If
        row_c = row_c + 1
    Loop
    
    
    Columns("E:G").Select
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("F2:F10000"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("G2:G10000"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("E2:E10000"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("E1:G10000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    row_c = 2
    
    Do While Cells(row_c, 5) <> ""
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(row_c, 5), Address:= _
        "http://ejje.weblio.jp/content/" & Cells(row_c, 5)
        If Cells(row_c, 7) = "" Then
            Cells(row_c, 7) = Now
        End If
        row_c = row_c + 1
    Loop


    Columns("I:K").Select
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("J2:J10000"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("K2:K10000"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("I2:I10000"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("I1:K10000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    row_c = 2
    
    Do While Cells(row_c, 9) <> ""
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(row_c, 9), Address:= _
        "http://ejje.weblio.jp/content/" & Cells(row_c, 9)
        If Cells(row_c, 11) = "" Then
            Cells(row_c, 11) = Now
        End If
        row_c = row_c + 1
    Loop


    Range("A1").Select
    
End Sub


Sub Color_hit_word()
'
' Macro1 Macro
'
    Dim i As Integer
    Dim iword As String

    iword = Cells(7, 4)
    Cells(7, 4).Interior.Color = 49407

    i = 2
    Do While Cells(i, 1) <> ""
        If Cells(i, 1) = iword Then
            Cells(i, 1).Interior.Color = 49407
        Else
           Cells(i, 1).Interior.Pattern = xlNone
        End If
        i = i + 1
    Loop
    

    i = 2
    Do While Cells(i, 5) <> ""
        If Cells(i, 5) = iword Then
            Cells(i, 5).Interior.Color = 49407
        Else
           Cells(i, 5).Interior.Pattern = xlNone
        End If
        i = i + 1
    Loop

    i = 2
    Do While Cells(i, 9) <> ""
        If Cells(i, 9) = iword Then
            Cells(i, 9).Interior.Color = 49407
        Else
           Cells(i, 9).Interior.Pattern = xlNone
        End If
        i = i + 1
    Loop

End Sub
