Sub auto_che_table1()
'
' Macro1 Macro
' �}�N���L�^�� : 2006/3/3  ���[�U�[�� : �֓��֑��Y
'

'
    Dim row As Integer
    Const COL = 1
    Dim hizuke As Date
    Dim wd As Integer
    Dim wdj As String
    Dim smonth As Integer
    Dim syear As Integer
        
    hizuke = InputBox("�N������͂��ĉ�����(��F2006/4)")
    
    row = 2
    smonth = Month(hizuke)
    syear = Year(hizuke)
    
    Cells(1, 1) = hizuke
    Cells(1, 1).NumberFormatLocal = "yyyy""�N""m""��"";@"
    
    Do While Month(hizuke) = smonth
        
    '�ϐ��u���t�l�v�̒l����Weekday�֐����g���ėj��������o��
        wd = Weekday(hizuke)

    '����o���ꂽ�j������{��ɕϊ�
        Select Case wd
            Case vbSunday
                wdj = "��"
            Case vbMonday
                wdj = "��"
            Case vbTuesday
                wdj = "��"
            Case vbWednesday
                wdj = "��"
            Case vbThursday
                wdj = "��"
            Case vbFriday
                wdj = "��"
            Case vbSaturday
                wdj = "�y"
        End Select
        
        Cells(row, COL) = hizuke
        Cells(row, COL).NumberFormatLocal = "m""��""d""��"";@"
        Cells(row, COL + 1) = wdj
        
        If wdj = "�y" Then
            Cells(row, COL + 1).Interior.ColorIndex = 34
        End If
        If wdj = "��" Then
            Cells(row, COL + 1).Interior.ColorIndex = 38
        End If
        hizuke = hizuke + 1
        row = row + 1
        
    Loop

    Range(Cells(1, 1), Cells(row - 1, 5)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    Range("A1:E1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Columns("A:A").Select
    Selection.ColumnWidth = 13
    Columns("B:B").Select
    Selection.ColumnWidth = 3
    Columns("C:C").Select
    Selection.ColumnWidth = 60
    Columns("D:D").Select
    Selection.ColumnWidth = 10
    Columns("E:E").Select
    Selection.ColumnWidth = 30
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "�\��"
    ActiveCell.Characters(1, 2).PhoneticCharacters = "���e�C"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "�؍ݏꏊ"
    ActiveCell.Characters(1, 2).PhoneticCharacters = "�^�C�U�C�o�V��"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "���l"
    ActiveCell.Characters(1, 2).PhoneticCharacters = "�r�R�E"
    
    Range("C1:E1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1").Select
    With Selection.Font
        .Name = "�l�r �o�S�V�b�N"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Selection.Font.Bold = True
    
End Sub
