Option Explicit
Sub Bt_ex()

    Dim w As Worksheet
    Dim ultima_cel As Range
    Dim resultado As Long
    resultado = 0
    Set w = Sheets(1)
    Set ultima_cel = w.Range("A1048576").End(xlUp)
    ultima_cel.Select
    
    
    Do While ActiveCell.Row >= 2
    
        resultado = resultado + ActiveCell.Value
        ActiveCell.Offset(-1, 0).Select
    Loop
    MsgBox "A soma é " & resultado
    Range("C2").Offset(1, 0).Value = resultado
    ultima_cel.Offset(1, 0).Value = "A soma dos numeros é " & resultado

End Sub
