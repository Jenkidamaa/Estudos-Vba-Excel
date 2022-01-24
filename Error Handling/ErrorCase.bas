Dim cmdDesbloquear_Click()

Dim PlanilhaAtual As Variant
Dim Senha As String
Dim w As worksheet

Senha = inputbox("Digite uma senha: ", vbOKonly)
ponto_saida:
  On Error Resume Next
  Set w = Nothing
  Exit Sub
erro_codigo:
  MsgBox "Planilha nao desbloqueadas. Senha inválida"
  Resume ponto_saida
  Exit Sub
    

