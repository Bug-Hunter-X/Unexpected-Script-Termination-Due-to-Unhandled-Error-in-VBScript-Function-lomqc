Function MyFunction(param1, param2)
  On Error GoTo ErrorHandler
  ' ... some code ...
  If param1 = "" Then
    Err.Raise 13, , "param1 cannot be empty"
  End If
  ' ... more code ...
  Exit Function
ErrorHandler:
  MsgBox "Error: " & Err.Description, vbCritical
  Err.Clear
End Function