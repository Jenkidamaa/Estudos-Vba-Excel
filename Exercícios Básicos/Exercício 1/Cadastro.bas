Attribute VB_Name = "Module1"
Sub Cadastro()
Application.ScreenUpdating = False
    
    Dim x As Integer
    x = Sheets(1).Select(xlUp) + x
    Sheets(1).Select
    Range("A2:F2").Copy
    Sheets(2).Select
    Range("A1048576").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    
    Sheets(1).Select
    Range("A2:F2").Delete
    
    Range("G5") = "Gravação ok"
    MsgBox "Processo Concluido...", vbOKOnly, "Concluido"
Application.ScreenUpdating = True

    
    
End Sub

Sub Cadastro2()
Application.ScreenUpdating = False
    'nome
    
    Sheets(1).Select
    Range("A2").Copy
    Sheets(2).Select
    Range("A1048576").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    
    'Endereço
    
    Sheets(1).Select
    Range("A5").Copy
    Sheets(2).Select
    Range("B1048576").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    
    'Bairro
        
    Sheets(1).Select
    Range("C2").Copy
    Sheets(2).Select
    Range("C1048576").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    
    'Cidade
        
    Sheets(1).Select
    Range("C5").Copy
    Sheets(2).Select
    Range("D1048576").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
        
    'CEP
        
    Sheets(1).Select
    Range("E2").Copy
    Sheets(2).Select
    Range("E1048576").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    
    'Telefone
        
    Sheets(1).Select
    Range("E5").Copy
    Sheets(2).Select
    Range("F1048576").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    
    Sheets(1).Select
    
    Range("A2").Clear
    Range("A5").Clear
    Range("C2").Clear
    Range("C5").Clear
    Range("E2").Clear
    Range("E5").Clear
    MsgBox "Cadastro realizado!", vbOKOnly, "Concluido"


Application.ScreenUpdating = True
        


End Sub
