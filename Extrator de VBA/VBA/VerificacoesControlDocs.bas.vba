Attribute VB_Name = "VerificacoesControlDocs"
Option Explicit

Public Function Verificar_NETFramework35() As Boolean
    
Dim WshShell As Object
Dim strKeyPath As String
Dim strValueName As String
Dim varValue As Variant
    
    ' Cria uma instância do WshShell para acessar o registro do Windows
    Set WshShell = CreateObject("WScript.Shell")
    
    ' Define o caminho da chave do registro para verificar a instalação do .NET Framework 3.5
    strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5\"
    strValueName = "Install"
    
    ' Tenta ler o valor da chave para verificar se o .NET Framework 3.5 está instalado
    On Error Resume Next ' Ignora erros se a chave não existir
    varValue = WshShell.RegRead(strKeyPath & strValueName)
    On Error GoTo 0 ' Desliga o tratamento de erro "On Error Resume Next"
    
    ' Limpa o objeto WshShell
    Set WshShell = Nothing
    
    ' Verifica se o .NET Framework 3.5 está instalado e ativo
    If varValue = 1 Then Verificar_NETFramework35 = True
    
End Function

Function VerificarConfiguracoesExcel() As Boolean

Dim WshShell As Object
Dim chaveMacros As String
Dim chaveConfiancaVBA As String
Dim valorMacros As Variant
Dim valorConfiancaVBA As Variant
    
    ' Cria uma instância do objeto Shell
    Set WshShell = CreateObject("WScript.Shell")
    
    ' Define o caminho da chave de registro para as configurações de macro
    chaveMacros = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Excel\Security\VBAWarnings"
    
    ' Define o caminho da chave de registro para a confiança no modelo de objeto do VBA
    chaveConfiancaVBA = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Excel\Security\AccessVBOM"
    
    ' Verifica as configurações de macro
    On Error Resume Next
    valorMacros = WshShell.RegRead(chaveMacros)
    valorConfiancaVBA = WshShell.RegRead(chaveConfiancaVBA)
    On Error GoTo 0
    
    ' Verifica se as configurações correspondem ao esperado (0 para desabilitado, 1 para habilitado)
    VerificarConfiguracoesExcel = (valorMacros = 1) And (valorConfiancaVBA = 1)
    
    ' Libera o objeto Shell
    Set WshShell = Nothing
    
End Function

Public Function VerificarConfiguracoesControlDocs() As Boolean
    
Dim vbResult As VbMsgBoxResult
Dim Msg As String
    
    If Not Verificar_NETFramework35 And Not VerificarConfiguracoesExcel Then
        
        Msg = "Você precisa fazer algumas configurações no seu computador e no Excel para que o ControlDocs funcione corretamente." & vbCrLf & vbCrLf
        Msg = Msg & "Clique no notão SIM para acessar o tutorial ensinando o passo a passo."
        
        vbResult = MsgBox(Msg, vbExclamation + vbYesNo, "Configuração do Computador")
        If vbResult = vbYes Then Call FuncoesLinks.AbrirUrl(urlConfigComputador)
        Exit Function
        
    Else
    
        If Not Verificar_NETFramework35 Then
            
            Msg = "Você precisa configurar o NET.Framework 3.5 para utilizar o ControlDocs." & vbCrLf & vbCrLf
            Msg = Msg & "Clique no notão SIM para acessar o tutorial ensinando o passo a passo."
            
            vbResult = MsgBox(Msg, vbExclamation + vbYesNo, "Configuração do Computador")
            If vbResult = vbYes Then Call FuncoesLinks.AbrirUrl(urlConfigComputador)
            Exit Function
            
        End If
        
        If Not VerificarConfiguracoesExcel Then
            
            Msg = "Você precisa configurar o Excel para utilizar o ControlDocs." & vbCrLf & vbCrLf
            Msg = Msg & "Clique no notão SIM para acessar o tutorial ensinando o passo a passo."
            
            vbResult = MsgBox(Msg, vbExclamation + vbYesNo, "Configuração do Computador")
            If vbResult = vbYes Then Call FuncoesLinks.AbrirUrl(urlConfigExcel)
            Exit Function
            
        End If
        
        VerificarConfiguracoesControlDocs = True
        Exit Function
    
    End If
    
End Function

Public Sub VerificarAtualizacao()
    
Dim vbResposta As VbMsgBoxResult
Dim BaseURL As String, Msg$
Dim RotaDownload As String
Dim VersaoAtual As String
Dim objXMLHTTP As Object
Dim VersaoNova As String
Dim timestamp As String
Dim objShell As Object
Dim json As Object
    
    timestamp = "?t=" & Format(Now, "yyyymmddhhnnss")
    BaseURL = "https://downloadcontroldocs.escoladaautomacaofiscal.com.br"
    
    VersaoAtual = Util.ApenasNumeros(ExtrairVersaoProjeto())
    
    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
    
    With objXMLHTTP
    
        .Open "GET", BaseURL & "/api/version" & timestamp, False
        
        .SetRequestHeader "Cache-Control", "no-cache, no-store, must-revalidate"
        .SetRequestHeader "Pragma", "no-cache"
        .SetRequestHeader "Expires", "0"
        
        On Error Resume Next
        .Send
        On Error GoTo 0
        
        If .Status <> 200 Then Exit Sub

    End With
    
    On Error Resume Next
        
        Set json = JsonConverter.ParseJson(objXMLHTTP.ResponseText)
        If Err.Number <> 0 Then Exit Sub
        
    On Error GoTo 0
    
    If json.Exists("latestVersion") And json.Exists("downloadUrl") Then
        
        VersaoNova = Util.ApenasNumeros(json("latestVersion"))
        RotaDownload = json("downloadUrl")
        
        If VersaoAtual < VersaoNova Then
            
            Msg = "Uma nova versão do ControlDocs está disponível para Download!" & vbCrLf & vbCrLf
            Msg = Msg & "Sua versão: " & FormatarVersao(ExtrairVersaoProjeto()) & vbCrLf
            Msg = Msg & "Nova versão: " & json("latestVersion") & vbCrLf & vbCrLf
            Msg = Msg & "Deseja baixar a nova versão agora?"
            
            vbResposta = MsgBox(Msg, vbQuestion + vbYesNo, "Atualização Disponível")
            
            If vbResposta = vbYes Then
                
                Set objShell = CreateObject("WScript.Shell")
                objShell.Run "cmd /c start " & BaseURL & RotaDownload, 1, False
                
            End If
            
        End If

    End If
    
    Set json = Nothing
    Set objXMLHTTP = Nothing
    Set objShell = Nothing
    
End Sub

Public Function FormatarVersao(versao As String) As String

Dim versaoLimpa As String
Dim versaoFormatada As String
    
    On Error GoTo TratarErro
    
    'Remove o "v" inicial se existir
    If Left(versao, 1) = "v" Then
        versaoLimpa = Mid(versao, 2)
    Else
        versaoLimpa = versao
    End If
    
    If Len(versaoLimpa) < 12 Then
        FormatarVersao = versao
        Exit Function
    End If
        
    Dim major As String
    Dim Ano As String
    Dim Mes As String
    Dim dia As String
    Dim build As String
    
    major = Left(versaoLimpa, 1)
    Ano = Mid(versaoLimpa, 2, 4)
    Mes = Mid(versaoLimpa, 6, 2)
    dia = Mid(versaoLimpa, 8, 2)
    build = Mid(versaoLimpa, 10)
    
    versaoFormatada = "v" & major & "." & Ano & "." & Mes & "." & dia & "." & build
    
    FormatarVersao = versaoFormatada
    Exit Function
    
TratarErro:
    
    FormatarVersao = versao
    
End Function
