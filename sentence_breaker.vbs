Sub SentenceBreaker()

    Dim somethingelse() As String
    somethingelse = Split(Range("B1").Value, " ")
    Dim num1, num2, num3 As Integer
    num1 = Cells(4, 1).Value
    num2 = Range("A5").Value
    num3 = Range("A6").Value
    MsgBox (somethingelse(3))

 End Sub
