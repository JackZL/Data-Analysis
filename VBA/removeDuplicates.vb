Sub removeNegative()
# delete the positive and negative value in column 18 if the column 2 and column 3 text are same
Dim array1() As Variant
Dim newsheet As Worksheet
Dim int0 As Integer
Dim int1 As Integer
Dim int2 As Integer
Dim cellFound1 As range
Dim cellFound2 As range
Dim cellPair As range
Dim string1 As String
Dim ce As range

'Application.ScreenUpdating = False
With ActiveSheet
    range("a1").CurrentRegion.AutoFilter 18, "-*"
    range(range("r1"), range("r1").End(xlDown)).Copy range("v1")
    'range("v:v").RemoveDuplicates 1, xlYes
    
    array1 = range(range("v2"), range("v2").End(xlDown))
    
    For int1 = 1 To UBound(array1, 1)
        string1 = Mid(array1(int1, 1), 2)
        range("a1").CurrentRegion.AutoFilter 18, "*" & string1
        Set cellFound1 = range("R:R").Find(array1(int1, 1), LookAt:=xlWhole)
        Set cellFound2 = range("R:R").Find(string1, LookAt:=xlWhole)
        If Not cellFound2 Is Nothing And Not cellFound1 Is Nothing Then
            If Cells(cellFound1.Row, 2) = Cells(cellFound2.Row, 2) And Cells(cellFound1.Row, 3) = Cells(cellFound2.Row, 3) Then
            Set cellPair = Union(cellFound1, cellFound2)
            cellPair.Interior.ColorIndex = 3
'                If cellFound1.Row < cellFound2.Row Then
'                    Rows(cellFound2.Row).Delete
'                    Rows(cellFound1.Row).Delete
'                Else
'                    Rows(cellFound1.Row).Delete
'                    Rows(cellFound2.Row).Delete
'                End If
            End If
        Else
            GoTo ff
        End If
ff:
    Next

End With
Application.ScreenUpdating = True
End Sub
