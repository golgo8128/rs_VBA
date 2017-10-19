Sub Color_focused_MDM2_intrx_Nicho()
'

    Dim i As Integer
    Dim k As Integer

    Dim enz(12) As String

    enz(1) = "ACAT1"
    enz(2) = "ACO2"
    enz(3) = "NME2"
    enz(4) = "ACLY"
    enz(5) = "CS"
    enz(6) = "ALDOA"
    enz(7) = "SHMT2"
    enz(8) = "MCCC2"
    enz(9) = "HSD17B4"
    enz(10) = "HADHA"
    enz(11) = "HADH"
    enz(12) = "HSD17B10"

    For i = 2 To 1967
        For k = 1 To 12
            If InStr(Cells(i, 2), enz(k)) <> 0 Then
                With Cells(i, 2).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            ElseIf InStr(Cells(i, 9), enz(k)) <> 0 Then
                With Cells(i, 9).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        Next k
    
    Next i
    
End Sub

Sub Color_focused_MDM2_intrx_BioGRID()
'
' Color_focused1 Macro
'

'

    Dim i As Integer
    Dim k As Integer

    Dim enz(12) As String

    enz(1) = "ACAT1"
    enz(2) = "ACO2"
    enz(3) = "NME2"
    enz(4) = "ACLY"
    enz(5) = "CS"
    enz(6) = "ALDOA"
    enz(7) = "SHMT2"
    enz(8) = "MCCC2"
    enz(9) = "HSD17B4"
    enz(10) = "HADHA"
    enz(11) = "HADH"
    enz(12) = "HSD17B10"

    For i = 2 To 1967
        For k = 1 To 12
            If Cells(i, 8) = enz(k) Or Cells(i, 9) = enz(k) Then
                With Cells(i, 8).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With Cells(i, 9).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        Next k
    
    Next i
    
End Sub
