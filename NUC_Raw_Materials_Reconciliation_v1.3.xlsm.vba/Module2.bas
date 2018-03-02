Sub Wrong()
  Dim i As Integer, a As Integer, b As Integer
  On Error GoTo Ignore
  For i = 1 To 2
    a = i / b
Ignore:
  Next
End Sub