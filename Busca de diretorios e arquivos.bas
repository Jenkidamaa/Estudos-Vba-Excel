Attribute VB_Name = "Module1"
Option Explicit
Private Sub btExecuta_click()

Dim W As Worksheet 'Manipula as planilhas
Dim vNome As String ' Nome do arquiv
Dim vArq As String
Dim vcaminho As String  'Caminho onde vamos salvar / apagar pastas
Dim vExisteArq As String 'Verificar se há arquivos
Application.ScreenUpdating = False

vcaminho = "C:\temp\teste\" 'Altera esse caminho para a pasta raiz desejada

Application.ScreenUpdating = True


vExisteArq = Dir(vcaminho & "*.*") '*.* verifica os arquivos da pasta

If vExisteArq <> "" Then

    Kill vcaminho & "*.*" 'Elimina todos os arquivos *.*
    RmDir vcaminho 'Remove diretorio
    Dir vcaminho 'Checar se há arquivos. Liberar a pasta
    
End If

'Recriar Diretorio
'----------------------------------------------------------



MkDir vcaminho

Application.DisplayAlerts = False

For Each W In Sheets

    vNome = W.Name    'Captura o nome da planilha
    'Verifica se o arquivo existe
    '-------------------------------------------------------
    vArq = Dir(vcaminho & vNome & ".xlsx")
    
    If vArq = "" Then
        Application.StatusBar = "Arquivo" & vNome & "não existe na pasta"
    End If
        
    W.Select
    W.Copy
    
    ActiveWorkbook.SaveAs Filename:=vcaminho & vNome,  FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close 'Fecha a nova pasta de trabalho
    
    


Next W

Application.DisplayAlerts = True


End Sub
