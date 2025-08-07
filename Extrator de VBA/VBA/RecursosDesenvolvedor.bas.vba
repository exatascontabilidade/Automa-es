Attribute VB_Name = "RecursosDesenvolvedor"
Option Explicit

Public Sub ComandosDesenvolvedor()
Attribute ComandosDesenvolvedor.VB_ProcData.VB_Invoke_Func = "E\n14"

Dim Comando As String
    
    Comando = InputBox("Digite o comando que deseja acessar", "Comandos de Desenvolvedor")
    Select Case True
        
        Case VBA.UCase(Comando) = "DATARESET"
            Call ResetarDados
        
        Case VBA.UCase(Comando) = "DELETARSPED"
            Call DeletarSPED
            
        Case VBA.UCase(Comando) = "PROTEGERREGISTROS"
            Call ProtegerRegistros
            
        Case VBA.UCase(Comando) = "DESPROTEGERREGISTROS"
            Call DesProtegerRegistros
        
        Case VBA.UCase(Comando) = "TESTE"
'            Call FuncoesSPEDFiscal.ImportarRegistrosSPEDFiscal(ActiveSheet.name)
        
        Case Else
            Call MsgBox("O comando informado não existe!", vbCritical, "Comando inválido")
            
    End Select

End Sub

Private Function ResetarDados()
    
Dim pasta As New Workbook
Dim Plan As Worksheet
            
    Set pasta = ThisWorkbook
    For Each Plan In pasta.Sheets
        
        Select Case True
            
            Case (Plan.CodeName <> "Autenticacao") And (Plan.CodeName <> "CadContrib")
                Call Util.DeletarDados(Plan, 4, False)
                
        End Select
    
    Next Plan
                
    MsgBox "Dados resetados com sucesso", vbInformation, "Reset de Dados"
    
End Function

Private Function DeletarSPED()
    
Dim pasta As New Workbook
Dim Plan As Worksheet
        
    Util.DesabilitarControles
        
        Set pasta = ThisWorkbook
        For Each Plan In pasta.Sheets
            
            Select Case True
                
                Case (Plan.CodeName <> "Autenticacao") And (Plan.CodeName <> "CadContrib") And InStr(1, Plan.CodeName, "reg")
                    Call Util.DeletarDados(Plan, 4, False)
                    
            End Select
        
        Next Plan
    
    Util.HabilitarControles
    
    MsgBox "Dados resetados com sucesso", vbInformation, "Reset de Dados"
    
End Function

Private Function DesProtegerRegistros()
    
Dim pasta As New Workbook
Dim Plan As Worksheet
    
    Call Util.DesabilitarControles
    
        Set pasta = ThisWorkbook
        For Each Plan In pasta.Sheets
            
            Select Case True
                
                Case InStr(1, Plan.CodeName, "reg")
'                    Call FuncoesPlanilha.DesprotegerPlanilha(Plan)
                    
            End Select
        
        Next Plan
                    
        MsgBox "Dados desprotegidos com sucesso!", vbInformation, "Desproteger Dados"
    
    Call Util.HabilitarControles
    
End Function

Private Function ProtegerRegistros()
    
Dim pasta As New Workbook
Dim Plan As Worksheet
            
    Call Util.DesabilitarControles
                
        Set pasta = ThisWorkbook
        For Each Plan In pasta.Sheets
            
            Select Case True
                
                Case VBA.Left(Plan.CodeName, 3) = "reg"
'                    Call FuncoesPlanilha.ProtegerPlanilha(Plan)
                    
            End Select
        
        Next Plan
                    
        MsgBox "Dados protegidos com sucesso!", vbInformation, "Proteger Dados"
    
    Call Util.HabilitarControles
        
End Function

Public Sub ExportarCodigoControlDocs(control As IRibbonControl)

Dim Projeto As VBProject
Dim componente As VBComponent
Dim caminhoPasta As String
Dim caminhoInicial As String
    
    ' Defina aqui o caminho padrão que você deseja usar
    caminhoInicial = "E:\Projetos\ControlDocs\"
    
    ' Solicitar ao usuário que selecione a pasta de destino
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = caminhoInicial
        .Title = "Selecione a pasta para exportar os módulos"
        .ButtonName = "Selecionar"
        If .Show = -1 Then ' Se uma pasta foi selecionada
            caminhoPasta = .SelectedItems(1) & "\"
        Else
            MsgBox "Nenhuma pasta selecionada. A exportação foi cancelada.", vbExclamation, "Exportação de Códigos do ControlDocs"
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
    
    MsgBox "Módulos exportados com sucesso para: " & caminhoPasta, vbInformation, "Exportação de Códigos do ControlDocs"
    
End Sub

Public Sub ImportarCodigoControlDocs(control As IRibbonControl)

Dim i As Long
Dim ARQUIVO As String
Dim Projeto As VBProject
Dim caminhoPasta As String
Dim caminhoInicial As String
Dim componente As VBComponent

    ' Defina aqui o caminho padrão que você deseja usar
    caminhoInicial = "E:\Projetos\ControlDocs\"
        
    ' Solicitar ao usuário que selecione a pasta de origem
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = caminhoInicial
        .Title = "Selecione a pasta para importar os módulos"
        .ButtonName = "Selecionar"
        If .Show = -1 Then ' Se uma pasta foi selecionada
            caminhoPasta = .SelectedItems(1) & "\"
        Else
            MsgBox "Nenhuma pasta selecionada. A importação foi cancelada.", vbExclamation, "Importação de Códigos do ControlDocs"
            Exit Sub
        End If
    End With
    
    ' Verificar se existem arquivos .bas ou .cls na pasta selecionada
    If Dir(caminhoPasta & "*.bas") = "" And Dir(caminhoPasta & "*.cls") = "" Then
        MsgBox "Nenhum arquivo .bas ou .cls encontrado na pasta selecionada.", vbExclamation, "Importação de Códigos do ControlDocs"
        Exit Sub
    End If
    
    ' Obter referência ao projeto VBA atual
    Set Projeto = ThisWorkbook.VBProject
    
    ' Remover todos os módulos existentes (exceto ThisWorkbook e RecursosDesenvolvedor)
    For i = Projeto.VBComponents.Count To 1 Step -1
        Set componente = Projeto.VBComponents(i)
        If componente.Type = vbext_ct_StdModule Or componente.Type = vbext_ct_ClassModule Then
            If componente.name <> "RecursosDesenvolvedor" Then
                Projeto.VBComponents.Remove componente
            End If
        End If
    Next i

    ' Percorrer todos os arquivos .bas na pasta
    ARQUIVO = Dir(caminhoPasta & "*.bas")
    Do While ARQUIVO <> ""
        If ARQUIVO <> "RecursosDesenvolvedor.bas" Then Projeto.VBComponents.Import caminhoPasta & ARQUIVO
        ARQUIVO = Dir()
    Loop

    ' Percorrer todos os arquivos .cls na pasta
    ARQUIVO = Dir(caminhoPasta & "*.cls")
    Do While ARQUIVO <> ""
        Projeto.VBComponents.Import caminhoPasta & ARQUIVO
        ARQUIVO = Dir()
    Loop
    
    MsgBox "Módulos importados com sucesso de: " & caminhoPasta, vbInformation, "Importação de Códigos do ControlDocs"
    
End Sub

Public Sub ExportarClasesModulosTXT() 'control As IRibbonControl)

Dim Projeto As VBProject
Dim componente As VBComponent
Dim caminhoPasta As String
Dim caminhoInicial As String
Dim arquivoModulos As String
Dim arquivoClasses As String
Dim conteudoModulos As String
Dim conteudoClasses As String
    
    ' Defina aqui o caminho padrão que você deseja usar
    caminhoInicial = "E:\Projetos\ControlDocs\"
    
    ' Solicitar ao usuário que selecione a pasta de destino
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .InitialFileName = caminhoInicial
        .Title = "Selecione a pasta para exportar os módulos"
        .ButtonName = "Selecionar"
        If .Show = -1 Then ' Se uma pasta foi selecionada
            caminhoPasta = .SelectedItems(1) & "\"
        Else
            MsgBox "Nenhuma pasta selecionada. A exportação foi cancelada.", vbExclamation, "Exportação de Códigos do ControlDocs"
            Exit Sub
        End If
    
    End With
    
    ' Definir nomes dos arquivos de saída
    arquivoModulos = caminhoPasta & "ModulosControlDocs.txt"
    arquivoClasses = caminhoPasta & "ClassesControlDocs.txt"
    
    ' Obter referência ao projeto VBA atual
    Set Projeto = ThisWorkbook.VBProject
    
    ' Percorrer todos os componentes do projeto
    For Each componente In Projeto.VBComponents
    
        ' Exportar módulos .bas
        If componente.Type = vbext_ct_StdModule Then
            
            conteudoModulos = conteudoModulos & "' --- " & componente.name & " ---" & vbNewLine
            conteudoModulos = conteudoModulos & componente.CodeModule.Lines(1, componente.CodeModule.CountOfLines) & vbNewLine & vbNewLine
        
        ' Exportar módulos de classe .cls
        ElseIf componente.Type = vbext_ct_ClassModule Then
            
            conteudoClasses = conteudoClasses & "' --- " & componente.name & " ---" & vbNewLine
            conteudoClasses = conteudoClasses & componente.CodeModule.Lines(1, componente.CodeModule.CountOfLines) & vbNewLine & vbNewLine
        
        End If
    
    Next componente
    
    ' Escrever conteúdo nos arquivos
    WriteToFile arquivoModulos, conteudoModulos
    WriteToFile arquivoClasses, conteudoClasses
    
    MsgBox "Módulos e Classes exportados com sucesso para: " & caminhoPasta, vbInformation, "Exportação de Códigos do ControlDocs"
    
End Sub

Private Sub WriteToFile(filePath As String, content As String)
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    
        Print #fileNum, content
        
    Close #fileNum
    
End Sub

Public Sub ExportarClasesModulosTXT2() 'control As IRibbonControl)

Dim Projeto As VBProject
Dim componente As VBComponent
Dim caminhoPasta As String
Dim caminhoInicial As String
Dim conteudoModulos As String
Dim conteudoClasses As String
Dim contadorLinhasModulos As Long
Dim contadorLinhasClasses As Long
Dim contadorArquivosModulos As Long
Dim contadorArquivosClasses As Long

    ' Defina aqui o caminho padrão que você deseja usar
    caminhoInicial = "G:\Meu Drive\EAF\OBSIDIAN\SegundoCerebro\1. PROJETOS\ControlDocs\Código Fonte"
    
    ' Solicitar ao usuário que selecione a pasta de destino
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = caminhoInicial
        .Title = "Selecione a pasta para exportar os módulos"
        .ButtonName = "Selecionar"
        If .Show = -1 Then ' Se uma pasta foi selecionada
            caminhoPasta = .SelectedItems(1) & "\"
        Else
            MsgBox "Nenhuma pasta selecionada. A exportação foi cancelada.", vbExclamation, "Exportação de Códigos do ControlDocs"
            Exit Sub
        End If
    End With
    
    ' Inicializar contadores
    contadorLinhasModulos = 0
    contadorLinhasClasses = 0
    contadorArquivosModulos = 1
    contadorArquivosClasses = 1
    
    ' Obter referência ao projeto VBA atual
    Set Projeto = ThisWorkbook.VBProject
    
    ' Percorrer todos os componentes do projeto
    For Each componente In Projeto.VBComponents
        Dim conteudoComponente As String
        Dim linhasComponente As Long
        
        If Not componente.CodeModule Is Nothing Then
        
            Dim linhasTotal As Long
            linhasTotal = componente.CodeModule.CountOfLines
            
            If linhasTotal > 0 Then
                conteudoComponente = "' --- " & componente.name & " ---" & vbNewLine & _
                                     componente.CodeModule.Lines(1, linhasTotal) & vbNewLine & vbNewLine
            Else
                conteudoComponente = "' --- " & componente.name & " --- (Vazio)" & vbNewLine & vbNewLine
            End If
            
        Else
            
            conteudoComponente = "' --- " & componente.name & " --- (Sem módulo de código)" & vbNewLine & vbNewLine
        
        End If
        
        linhasComponente = UBound(Split(conteudoComponente, vbNewLine)) + 1
        
        ' Exportar módulos .bas
        If componente.Type = vbext_ct_StdModule Then
            If contadorLinhasModulos + linhasComponente > 10000 Then
                ' Escrever conteúdo atual no arquivo e iniciar um novo
                WriteToFile caminhoPasta & "ModulosControlDocs_" & contadorArquivosModulos & ".txt", conteudoModulos
                conteudoModulos = ""
                contadorLinhasModulos = 0
                contadorArquivosModulos = contadorArquivosModulos + 1
            End If
            conteudoModulos = conteudoModulos & conteudoComponente
            contadorLinhasModulos = contadorLinhasModulos + linhasComponente
        ' Exportar módulos de classe .cls
        ElseIf componente.Type = vbext_ct_ClassModule Then
            If contadorLinhasClasses + linhasComponente > 10000 Then
                ' Escrever conteúdo atual no arquivo e iniciar um novo
                WriteToFile caminhoPasta & "ClassesControlDocs_" & contadorArquivosClasses & ".txt", conteudoClasses
                conteudoClasses = ""
                contadorLinhasClasses = 0
                contadorArquivosClasses = contadorArquivosClasses + 1
            End If
            conteudoClasses = conteudoClasses & conteudoComponente
            contadorLinhasClasses = contadorLinhasClasses + linhasComponente
        End If
    Next componente
    
    ' Escrever o conteúdo restante nos arquivos finais
    If conteudoModulos <> "" Then
        WriteToFile caminhoPasta & "ModulosControlDocs_" & contadorArquivosModulos & ".txt", conteudoModulos
    End If
    If conteudoClasses <> "" Then
        WriteToFile caminhoPasta & "ClassesControlDocs_" & contadorArquivosClasses & ".txt", conteudoClasses
    End If
    
    MsgBox "Módulos e Classes exportados com sucesso para: " & caminhoPasta & vbNewLine & _
           "Arquivos de Módulos: " & contadorArquivosModulos & vbNewLine & _
           "Arquivos de Classes: " & contadorArquivosClasses, _
           vbInformation, "Exportação de Códigos do ControlDocs"
    
End Sub

Public Sub ExportarClasesModulosTXT3() 'control As IRibbonControl)
    Dim Projeto As VBProject
    Dim componente As VBComponent
    Dim caminhoPasta As String
    Dim caminhoInicial As String
    Dim conteudoModulos As String
    Dim conteudoClasses As String
    Dim contadorCaracteresModulos As Long
    Dim contadorCaracteresClasses As Long
    Dim contadorArquivosModulos As Long
    Dim contadorArquivosClasses As Long
    Const LIMITE_CARACTERES As Long = 350000
    
    ' Defina aqui o caminho padrão que você deseja usar
    caminhoInicial = "G:\Meu Drive\EAF\OBSIDIAN\SegundoCerebro\1. PROJETOS\ControlDocs\Código Fonte"
    
    ' Solicitar ao usuário que selecione a pasta de destino
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = caminhoInicial
        .Title = "Selecione a pasta para exportar os módulos"
        .ButtonName = "Selecionar"
        If .Show = -1 Then ' Se uma pasta foi selecionada
            caminhoPasta = .SelectedItems(1) & "\"
        Else
            MsgBox "Nenhuma pasta selecionada. A exportação foi cancelada.", vbExclamation, "Exportação de Códigos do ControlDocs"
            Exit Sub
        End If
    End With
    
    ' Inicializar contadores
    contadorCaracteresModulos = 0
    contadorCaracteresClasses = 0
    contadorArquivosModulos = 1
    contadorArquivosClasses = 1
    
    ' Obter referência ao projeto VBA atual
    Set Projeto = ThisWorkbook.VBProject
    
    ' Percorrer todos os componentes do projeto
    For Each componente In Projeto.VBComponents
        Dim conteudoComponente As String
        Dim caracteresComponente As Long
        
        conteudoComponente = ObterConteudoComponente(componente)
        caracteresComponente = Len(conteudoComponente)
        
        ' Exportar módulos .bas
        If componente.Type = vbext_ct_StdModule Then
            If contadorCaracteresModulos + caracteresComponente > LIMITE_CARACTERES Then
                ' Escrever conteúdo atual no arquivo e iniciar um novo
                WriteToFile caminhoPasta & "ModulosControlDocs_" & contadorArquivosModulos & ".txt", conteudoModulos
                conteudoModulos = ""
                contadorCaracteresModulos = 0
                contadorArquivosModulos = contadorArquivosModulos + 1
            End If
            conteudoModulos = conteudoModulos & conteudoComponente
            contadorCaracteresModulos = contadorCaracteresModulos + caracteresComponente
        ' Exportar módulos de classe .cls
        ElseIf componente.Type = vbext_ct_ClassModule Then
            If contadorCaracteresClasses + caracteresComponente > LIMITE_CARACTERES Then
                ' Escrever conteúdo atual no arquivo e iniciar um novo
                WriteToFile caminhoPasta & "ClassesControlDocs_" & contadorArquivosClasses & ".txt", conteudoClasses
                conteudoClasses = ""
                contadorCaracteresClasses = 0
                contadorArquivosClasses = contadorArquivosClasses + 1
            End If
            conteudoClasses = conteudoClasses & conteudoComponente
            contadorCaracteresClasses = contadorCaracteresClasses + caracteresComponente
        End If
    Next componente
    
    ' Escrever o conteúdo restante nos arquivos finais
    If conteudoModulos <> "" Then
        WriteToFile caminhoPasta & "ModulosControlDocs_" & contadorArquivosModulos & ".txt", conteudoModulos
    End If
    If conteudoClasses <> "" Then
        WriteToFile caminhoPasta & "ClassesControlDocs_" & contadorArquivosClasses & ".txt", conteudoClasses
    End If
    
    MsgBox "Módulos e Classes exportados com sucesso para: " & caminhoPasta & vbNewLine & _
           "Arquivos de Módulos: " & contadorArquivosModulos & vbNewLine & _
           "Arquivos de Classes: " & contadorArquivosClasses, _
           vbInformation, "Exportação de Códigos do ControlDocs"
    
End Sub

Private Function ObterConteudoComponente(componente As VBComponent) As String
    Dim conteudoComponente As String
    Dim linhasComponente As Long
    
    If Not componente.CodeModule Is Nothing Then
        Dim linhasTotal As Long
        linhasTotal = componente.CodeModule.CountOfLines
        
        If linhasTotal > 0 Then
            conteudoComponente = "' --- " & componente.name & " ---" & vbNewLine & _
                                 componente.CodeModule.Lines(1, linhasTotal) & vbNewLine & vbNewLine
        Else
            conteudoComponente = "' --- " & componente.name & " --- (Vazio)" & vbNewLine & vbNewLine
        End If
    Else
        conteudoComponente = "' --- " & componente.name & " --- (Sem módulo de código)" & vbNewLine & vbNewLine
    End If

    linhasComponente = UBound(Split(conteudoComponente, vbNewLine)) + 1
    
    ObterConteudoComponente = conteudoComponente
End Function


