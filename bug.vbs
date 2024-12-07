Function MyFunction(param1)
  ' Some code here
  If param1 = "" Then
    Err.Raise 13, , "Parameter cannot be empty"
  End If
  ' More code here
End Function