Sub Signif_Cell_Color1_2()
'
' Signif_Cell_Color1_2 Macro
'
    Dim i As Integer
    Dim j As Integer
    
    For i = 2 To 500
        For j = 14 To 35 Step 2
    
        If Cells(i, j) <> "" And Cells(i, j).Value < 0.05 Then
            Cells(i, j).Interior.ColorIndex = 6
        End If
    
        Next
        
        For j = 15 To 36 Step 2
        
        If Cells(i, j) <> "" And Cells(i, j).Value < 0.05 Then
            Cells(i, j).Interior.ColorIndex = 42
        End If
        
        Next
        
    Next
    
End Sub

