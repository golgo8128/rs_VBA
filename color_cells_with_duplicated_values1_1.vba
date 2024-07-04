' Creation of this VBA started with ChatGPT 3.5
Sub ColorCellsWithDuplicatedValues()

    Dim sel_range As Range
    Dim cur_cell As Range
    Dim val_to_count_h As Object
    Dim val_to_coloridx_h As Object
    Dim cur_coloridx As Integer
    
    Dim curVal As Variant
    
    
    ' Define the range where you want to check for duplicates
    Set sel_range = Selection ' Change this to your desired range
    
    ' Dictionary object to store counts of each value
    Set val_to_count_h = CreateObject("Scripting.Dictionary")
    Set val_to_coloridx_h = CreateObject("Scripting.Dictionary")
    
    ' Loop through each cell in the range
    For Each cur_cell In sel_range
        If cur_cell.Value <> "" Then ' Check if the cell is not empty
            If val_to_count_h.exists(cur_cell.Value) Then
                ' If value already exists in dictionary, color the cell
                val_to_count_h(cur_cell.Value) = val_to_count_h(cur_cell.Value) + 1
                'cur_cell.Interior.Color = RGB(255, 192, 0) ' Change the color as needed
            Else
                ' Add value to dictionary
                val_to_count_h(cur_cell.Value) = 1
            End If
        End If
    Next cur_cell
    
    cur_coloridx = 3
    For Each curVal In val_to_count_h.Keys
        If val_to_count_h(curVal) > 1 Then
            val_to_coloridx_h(curVal) = cur_coloridx
            cur_coloridx = cur_coloridx + 1
            If cur_coloridx > 56 Then
                cur_coloridx = 3
            End If
        End If
    Next curVal
    
    For Each cur_cell In sel_range
        If val_to_coloridx_h.exists(cur_cell.Value) Then
                cur_cell.Interior.ColorIndex = val_to_coloridx_h(cur_cell.Value)
        End If
    Next cur_cell
    
    
    ' Clean up
    Set val_to_count_h = Nothing
    Set val_to_coloridx_h = Nothing
    
End Sub
