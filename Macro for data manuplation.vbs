'reduce duplicate
'x is number of row, y is number of column, k is preset value of column searching quanitty
Sub macro()
k = 0
For j = 1 To 100
    If Cells(1, 2 * j) <> 0 Then
        k = k + 1
    End If
Next j
For j = 1 To k
    po = 0
    a = 0
    For i = 1 To 700
        If Cells(i, 2 * j).Value <> 0 Then
            a = a + 1
        End If
    Next i

    'a is number of lines
    'b is number of duplicated cells
    
    'sum
   
    For i = 1 To a
        Sum = Sum + Cells(i, j * 2).Value
        SumT = SumT + Cells(i, j * 2 - 1).Value
        b = b + 1
        If Cells(i, (j - 1) * 2 + 1).Value <> Cells(i + 1, (j - 1) * 2 + 1).Value Then
            Average = Sum / b
            AverageT = SumT / b
            po = po + 1
            Cells(po, 2 * k + j * 2).Value = Average
            Cells(po, 2 * k + j * 2 - 1).Value = AverageT
            resi = (2 * k + j * 2 - 1) Mod 26
            If 2 * k + j * 2 - 1 > 26 And 2 * k + j * 2 - 1 < 52 Then
                Range(Chr(65) & Chr(64 + resi) & po).Interior.ColorIndex = 4
            ElseIf 2 * k + j * 2 - 1 > 52 And 2 * k + j * 2 - 1 <= 78 Then
                Range(Chr(66) & Chr(64 + resi) & po).Interior.ColorIndex = 4
            ElseIf 2 * k + j * 2 - 1 < 26 Then
                Range(Chr(64 + 2 * k + j * 2 - 1) & po).Interior.ColorIndex = 4
            End If
            b = 0
            Sum = 0
            SumT = 0
            
        End If
    Next i
Next j
For i = 1 To po
    a = 0
    For j = k + 1 To 2 * k
        a = a + Cells(i, 2 * j).Value
    Next j
    Cells(i, 4 * k + 1).Value = -a / k    
Next i
End Sub