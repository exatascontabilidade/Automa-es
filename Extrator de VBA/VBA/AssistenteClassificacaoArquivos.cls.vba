Attribute VB_Name = "AssistenteClassificacaoArquivos"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public CaminhoPowerShell As String

Public Function ListarArquivos(ByVal CaminhoPastaRaiz As String, Optional Extensao As String = "*.*") As Variant

Dim ListaPastas As Variant, ListaArquivos
Dim ConteudoPaths As String, CaminhoArquivoPaths$, CaminhoArquivoSaida$, CaminhoScript$, FiltroExtensao$
    
    On Error GoTo Notificar:
    
    'Obter todas as pastas e subpastas (incluindo a raiz)
    ListaPastas = ListarPastas(CaminhoPastaRaiz)
    
    'Preparar filtro de extensão
    FiltroExtensao = Extensao
    If FiltroExtensao <> "*.*" And Left(FiltroExtensao, 1) <> "." Then FiltroExtensao = "." & FiltroExtensao
    
    'Montar conteúdo dos paths
    ConteudoPaths = VBA.Join(ListaPastas, vbCrLf)
    CaminhoArquivoPaths = CriarArquivoPaths(ConteudoPaths)
    
    'Criar arquivo temporário para saída
    CaminhoArquivoSaida = Replace(CaminhoArquivoPaths, ".tmp", "_arquivos.txt")
    
    'Criar script PowerShell
    CaminhoScript = CriarScriptPowerShell_ListarArquivos(CaminhoArquivoPaths, CaminhoArquivoSaida, FiltroExtensao)
    
    'Executar PowerShell
    Call ExecutarScriptPowerShell(CaminhoScript)
    
    'Extrair resultado como array
    ListaArquivos = Util.ExtrairArrayArquivoTXT(CaminhoArquivoSaida)
    
    'Limpar arquivos temporários
    Call DeletarArquivosTemporarios(Array(CaminhoArquivoSaida, CaminhoScript, CaminhoArquivoPaths))
    
    ListarArquivos = ListaArquivos
    
Exit Function
Notificar:

    Call TratarExcecoes
    
End Function

Public Function ListarPastas(ByVal CaminhoPastaRaiz As String) As Variant

Dim CaminhoArquivoPaths As String, CaminhoArquivoSaida$, CaminhoScript$
Dim LinhasScript As Variant, ListaPastas
    
    'Criar arquivo com o caminho da pasta raiz em UTF-8
    CaminhoArquivoPaths = CriarArquivoPaths(CaminhoPastaRaiz)
    
    'Criar arquivo temporário para saída
    CaminhoArquivoSaida = Replace(CaminhoArquivoPaths, ".tmp", "_pastas.txt")
    
    'Montar linhas do script
    CaminhoScript = CriarScriptPowerShell_ListarPastas(CaminhoArquivoPaths, CaminhoArquivoSaida)
    
    'Executar PowerShell
    Call ExecutarScriptPowerShell(CaminhoScript)
    
    'Extrai resultado como array
    ListaPastas = Util.ExtrairArrayArquivoTXT(CaminhoArquivoSaida)
    
    'Limpar arquivos temporários
    Call DeletarArquivosTemporarios(Array(CaminhoArquivoSaida, CaminhoScript, CaminhoArquivoPaths))
    
    ListarPastas = ListaPastas
    
End Function

Private Function CriarArquivoPaths(ByVal Conteudo As String) As String

Dim objFSO As New FileSystemObject
Dim ArqTemporario As String
Dim objStream As Object
    
    ArqTemporario = objFSO.GetSpecialFolder(2) & "\" & objFSO.GetTempName
    
    Set objStream = CreateObject("ADODB.Stream")
    
    With objStream
        
        .Type = 2
        .Charset = "windows-1252"
        .Open
        If VBA.Len(Conteudo) > 0 Then .WriteText Conteudo
        .SaveToFile ArqTemporario, 2
        
        .Close
        
    End With
    
    Set objStream = Nothing
    Set objFSO = Nothing
    
    CriarArquivoPaths = ArqTemporario
    
End Function

Private Function CriarScriptPowerShell_ListarPastas(ByVal CaminhoArquivoPaths As String, ByVal CaminhoArquivoSaida As String) As String

Dim objFSO As New FileSystemObject
Dim ARQUIVO As Object
Dim ScriptPath As String
Dim ScriptContent As String
    
    ScriptPath = objFSO.GetSpecialFolder(2) & "\" & objFSO.GetTempName & ".ps1"
    Set ARQUIVO = objFSO.CreateTextFile(ScriptPath, True)
    
    With ARQUIVO
        
        .WriteLine "# Ler o arquivo de paths com Windows-1252 (ANSI)" & vbCrLf
        .WriteLine "$path = Get-Content -Path '" & CaminhoArquivoPaths & "' -Encoding 'Default'" & vbCrLf
        .WriteLine "if (Test-Path -LiteralPath $path) {" & vbCrLf
        .WriteLine "    $content = @()" & vbCrLf
        .WriteLine "    $content += $path" & vbCrLf
        .WriteLine "    Get-ChildItem -LiteralPath $path -Directory -Recurse | ForEach-Object { $content += $_.FullName }" & vbCrLf
        .WriteLine "    # Usar [IO.File]::WriteAllLines com Windows-1252 (ANSI)" & vbCrLf
        .WriteLine "    [IO.File]::WriteAllLines('" & CaminhoArquivoSaida & "', $content, [System.Text.Encoding]::GetEncoding(1252))"
        .WriteLine "}" & vbCrLf

        .Close
        
    End With
    
    Set objFSO = Nothing
    Set ARQUIVO = Nothing
    
    CriarScriptPowerShell_ListarPastas = ScriptPath
    
End Function

Private Function CriarScriptPowerShell_ListarArquivos(ByVal CaminhoArquivoPaths As String, ByVal CaminhoArquivoSaida As String, ByVal FiltroExtensao As String) As String
    
Dim objFSO As New FileSystemObject
Dim Script As String
Dim ARQUIVO As Object
    
    Script = objFSO.GetSpecialFolder(2) & "\" & objFSO.GetTempName & ".ps1"
    
    Set ARQUIVO = objFSO.CreateTextFile(Script, True)
    With ARQUIVO
                
        .WriteLine "# Ler o arquivo de paths com Windows-1252 (ANSI)"
        .WriteLine "$paths = Get-Content -Path '" & CaminhoArquivoPaths & "' -Encoding 'Default'"
        .WriteLine "$results = @()"
        .WriteLine "$paths | ForEach-Object {"
        .WriteLine "    if (Test-Path -LiteralPath $_) {"
        .WriteLine "        $results += Get-ChildItem -LiteralPath $_ -Filter '*" & FiltroExtensao & "' -File -Recurse | Select-Object -ExpandProperty FullName"
        .WriteLine "    }"
        .WriteLine "}"
        .WriteLine "    # Usar [IO.File]::WriteAllLines com Windows-1252 (ANSI)" & vbCrLf
        .WriteLine "    [IO.File]::WriteAllLines('" & CaminhoArquivoSaida & "', $results, [System.Text.Encoding]::GetEncoding(1252))" & vbCrLf

        .Close
        
    End With
    
    CriarScriptPowerShell_ListarArquivos = Script
    
End Function

Private Sub ExecutarScriptPowerShell(ByVal CaminhoScript As String)
    
Dim objShell As Object
Dim Comando As String
    
    Set objShell = CreateObject("WScript.Shell")
    Comando = """" & CaminhoPowerShell & """ -NoProfile -ExecutionPolicy Bypass -File """ & CaminhoScript & """"
    
    objShell.Run Comando, 0, True
    objShell.Run "cmd /c ping -n 2 localhost >nul", 0, True

    Set objShell = Nothing
    
End Sub

Private Sub DeletarArquivosTemporarios(ByVal Arquivos As Variant)

Dim objFSO As New FileSystemObject
Dim ARQUIVO As Variant
    
    On Error Resume Next
        
        For Each ARQUIVO In Arquivos
            
            If objFSO.FileExists(ARQUIVO) Then objFSO.DeleteFile ARQUIVO
            
        Next
        
    On Error GoTo 0
    
End Sub

Public Function EncontrouPowerShell() As Boolean

Dim Caminhos As Variant, Caminho
    
    Caminhos = Array("C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe", _
        "C:\Program Files\PowerShell\7\pwsh.exe", "powershell.exe", "pwsh.exe")
        
    For Each Caminho In Caminhos
        
        If Dir(Caminho) <> "" Then
            
            CaminhoPowerShell = Caminho
            EncontrouPowerShell = True
            Exit Function
            
        End If
        
    Next
    
    Call MensagemAlertaPowerShell
    
End Function

Private Function TratarExcecoes()

    Select Case True
        
        Case Err.Number = -2147024671
            Call MensagemAlertaAntiVirus
            
        Case Else
            With infNotificacao
            
                .Funcao = "ListarArquivos"
                .Classe = "AssistenteClassificacaoArquivos"
                .MensagemErro = Err.Number & " - " & Err.Description
                .OBSERVACOES = "Erro Inesperado"
                
            End With
            
            Call Notificacoes.NotificarErroInesperado
            
    End Select

End Function

Private Function MensagemAlertaPowerShell()

Dim Msg As String
    
    Msg = "O PowerShell não foi detectado em seu computador." & vbCrLf & vbCrLf
    Msg = Msg & "Para utilizar esta funcionalidade, é necessário ter o PowerShell instalado." & vbCrLf
    Msg = Msg & "Se precisar de ajuda nesse processo, conte com o nosso suporte técnico."
    
    Call Util.MsgAlerta(Msg, "Power Shell Não Detectado")

End Function

Private Function MensagemAlertaAntiVirus()

Dim Msg As String

    Msg = "O ControlDocs foi bloqueado por alguma proteção do seu computador (antivírus ou Windows Defender)." & vbCrLf
    Msg = Msg & "Por favor, adicione o ControlDocs à lista de permissões do seu antivírus ou consulte o suporte técnico." & vbCrLf
    
    Call Util.MsgCritica(Msg, "Execução Bloqueada")
    
    With infNotificacao
    
        .Funcao = "ListarArquivos"
        .Classe = "AssistenteClassificacaoArquivos"
        .MensagemErro = Err.Number & " - " & Err.Description

    End With
    
    Call Notificacoes.EnviarNotificacaoErro
    
End Function


