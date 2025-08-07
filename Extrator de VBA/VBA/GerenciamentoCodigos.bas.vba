Attribute VB_Name = "GerenciamentoCodigos"
Option Explicit

Public Sub ExportarModulos()

Dim Projeto As VBProject
Dim componente As VBComponent
Dim caminhoPasta As String
    
    ' Solicitar ao usuário que selecione a pasta de destino
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Selecione a pasta para exportar os módulos"
        .ButtonName = "Selecionar"
        If .Show = -1 Then ' Se uma pasta foi selecionada
            caminhoPasta = .SelectedItems(1) & "\"
        Else
            MsgBox "Nenhuma pasta selecionada. A exportação foi cancelada."
            Exit Sub
        End If
    End With
    
    ' Obter referência ao projeto VBA atual
    Set Projeto = ThisWorkbook.VBProject
    
    ' Percorrer todos os componentes do projeto
    For Each componente In Projeto.VBComponents
        ' Exportar módulos .bas e módulos de classe .cls
        If componente.Type = vbext_ct_StdModule Or componente.Type = vbext_ct_ClassModule Then
            componente.Export caminhoPasta & componente.name & IIf(componente.Type = vbext_ct_StdModule, ".bas", ".cls")
        End If
    Next componente
    
    MsgBox "Módulos exportados com sucesso para: " & caminhoPasta
End Sub

Public Sub ImportarModulos()
    Dim Projeto As VBProject
    Dim caminhoPasta As String
    Dim ARQUIVO As String
    
    ' Solicitar ao usuário que selecione a pasta de origem
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Selecione a pasta para importar os módulos"
        .ButtonName = "Selecionar"
        If .Show = -1 Then ' Se uma pasta foi selecionada
            caminhoPasta = .SelectedItems(1) & "\"
        Else
            MsgBox "Nenhuma pasta selecionada. A importação foi cancelada."
            Exit Sub
        End If
    End With
    
    ' Verificar se existem arquivos .bas ou .cls na pasta selecionada
    If Dir(caminhoPasta & "*.bas") = "" And Dir(caminhoPasta & "*.cls") = "" Then
        MsgBox "Nenhum arquivo .bas ou .cls encontrado na pasta selecionada."
        Exit Sub
    End If
    
    ' Obter referência ao projeto VBA atual
    Set Projeto = ThisWorkbook.VBProject
    
    ' Percorrer todos os arquivos .bas na pasta
    ARQUIVO = Dir(caminhoPasta & "*.bas")
    Do While ARQUIVO <> ""
        Projeto.VBComponents.Import caminhoPasta & ARQUIVO
        ARQUIVO = Dir()
    Loop
    
    ' Percorrer todos os arquivos .cls na pasta
    ARQUIVO = Dir(caminhoPasta & "*.cls")
    Do While ARQUIVO <> ""
        Projeto.VBComponents.Import caminhoPasta & ARQUIVO
        ARQUIVO = Dir()
    Loop
    
    MsgBox "Módulos importados com sucesso de: " & caminhoPasta
End Sub

