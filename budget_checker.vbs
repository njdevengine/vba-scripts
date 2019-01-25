Sub budget()


Dim budget As Integer
budget = Range("C3").Value

Dim price As Integer
price = Range("F3").Value

Dim fee As Integer
fees = Range("H3").Value

Dim feeval As Integer
feeval = price * fees

Range("L3").Value = feeval + price

If (feeval + price < budget) Then
    MsgBox ("within budget")
Else: MsgBox ("over budget")
End If

End Sub
