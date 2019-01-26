Sub nestedloop()
Dim hornets As Integer
hornets = 0

  For r = 1 To 6

    For c = 1 To 7
         If Cells(r, c) = "Hornets" Then

            hornets = hornets + 1
            Cells(r, c).Value = "Bugs"
            MsgBox ("hornets is" + Str(hornets))
            End If

     Next c

  Next r

End Sub

