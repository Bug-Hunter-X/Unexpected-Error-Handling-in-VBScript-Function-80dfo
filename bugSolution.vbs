Function MyFunction(param1)
  On Error GoTo ErrHandler
  ' Some code here
  If param1 = "" Then
    Err.Raise 13, , "Parameter cannot be empty"
  End If
  ' More code here
  Exit Function
ErrHandler:
  ' Handle the error gracefully
  MsgBox "An error occurred: " & Err.Description, vbCritical
  Err.Clear
End Function