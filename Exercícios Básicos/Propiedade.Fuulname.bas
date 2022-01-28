
'Essa função retorna o caminho do disco e o nome da pasta de trabalho ativa.
Function RetornaCaminho() As String
  Let RetornaCaminho = ThisWorkbook.FullName
End Function

'A função retorna apenas o nome da pasta de trabalho.
Function RetornaNomeArq() As String
  Let RetornaNomeArq = ThisWorkbook.Name
End Function

'Esse exemplo de macro mostra caminho do arquivo ativo.
Sub teste()
  MsgBox ActiveWorkbook.FullName
End Sub
