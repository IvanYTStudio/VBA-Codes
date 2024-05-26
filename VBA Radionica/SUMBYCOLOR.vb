Function SUMBYCOLOR(sum_range As Range, color_cell As Range) As Double
Dim cell As Range
Application.Volatile

If color_cell.Cells.Count <> 1 Then
    SUMBYCOLOR = CVErr(xlErrValue)
End If

For Each cell In sum_range

    If cell.Interior.color = color_cell.Interior.color Then
        SUMBYCOLOR = SUMBYCOLOR + cell
    End If

Next cell

End Function
