Function FirstName()
  Dim FullName As String
  Dim FirstSpace As Integer
  FullName = Application.UserName
  FirstSpace = InStr(FullName, “ “)
  If FirstSpace = 0 then
    FirstName = FullName
  Else
    FirstName = Left(FullName, FirstSpace – 1)
  End If
End Function
