Function f(x)
  If IsNumeric(x) Then 
    If CDbl(x) < 0 Then
      f = -CDbl(x)
    Else
      f = CDbl(x)
    End If
  Else
    MsgBox "Invalid input. Please enter a number."
  End If
End Function

MsgBox f(-5) 