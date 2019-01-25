Sub lotto()

    For i = 2 To 1001

        If (Str(Cells(i, 3)) = "3957481") Then
            Cells(2, 6) = Cells(i, 1)
            Cells(2, 7) = Cells(i, 2)
            Cells(2, 8) = Cells(i, 3)
        ElseIf (Str(Cells(i, 3)) = "5865187") Then
            Cells(3, 6) = Cells(i, 1)
            Cells(3, 7) = Cells(i, 2)
            Cells(3, 8) = Cells(i, 3)
        ElseIf (Str(Cells(i, 3)) = "2817729") Then
            Cells(4, 6) = Cells(i, 1)
            Cells(4, 7) = Cells(i, 2)
            Cells(4, 8) = Cells(i, 3)
        ' ElseIf
        End If
    Next i
MsgBox ("Congratulations !!! " + Cells(2, 6) + Cells(2, 7))
End Sub
