Sub hello()
    MsgBox ("hello")
End Sub
Sub reset()
    Dim originX As Integer
    Dim originY As Integer
    originX = 2
    originY = 2

    For q = 1 To 6
        If q <= 4 Then
            Call onesurface(q * 10, originX + q * 3 - 3, originY + 3)
        ElseIf q = 5 Then
            Call onesurface(q * 10, originX + 3, originY)
        ElseIf q = 6 Then
            Call onesurface(q * 10, originX + 3, originY + 6)
        End If
    Next q
End Sub

Sub onesurface(qube As Integer, x As Integer, y As Integer)
    Dim count As Integer
    count = 1
    For i = 0 To 2
        For j = 0 To 2
            Cells(y + i, x + j).Value = qube + count
            count = count + 1
        Next j
    Next i
End Sub
