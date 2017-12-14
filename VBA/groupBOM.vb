Private Sub GroupBomLevel_Click()
    Dim RowsCount As Integer
    RowsCount = Range("A3").CurrentRegion.Rows.Count
    
    Level = 5 'set the level if you like to have more levels
    For i = 1 To Level
        Range("A3").Select
        For j = 1 To RowsCount - 2
            With Selection
            If .Value <= i Then
                .Offset(1, 0).Select
            Else
                .Rows.Group
                .Offset(1, 0).Select
            End If
            End With
        Next
    Next
    
    With ActiveSheet.Outline
        .AutomaticStyles = False
        .SummaryRow = xlAbove
        .SummaryColumn = xlRight
    End With
    'set color
    With Columns("A:A")
        .FormatConditions.AddColorScale ColorScaleType:=3
    End With
End Sub
