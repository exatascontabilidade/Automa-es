Attribute VB_Name = "ApiNotificacoes"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Constantes da API (ajuste conforme necessário)
Private Const API_URL As String = "https://telegram.escoladaautomacaofiscal.com.br/"
Private Const INTERVALO_MINIMO As Double = 5 ' Segundos entre mensagens

' Variável para controle de rate limiting
Private ultimoEnvio As Double

Public Function EnviarNotificacaoErro() As Boolean

Dim http As New XMLHTTP60
Dim MsgHTML As String
    
    On Error GoTo Sair
    
    MsgHTML = FormatarNotificacaoHTML()
    MsgHTML = EscapeJSON(MsgHTML)
    
    Dim corpoJSON As String
    corpoJSON = "{""mensagem"": """ & MsgHTML & """, " & """silencioso"": false, " & """prioridade"": ""alta""}"
    
    http.Open "POST", API_URL & "enviar-notificacao", False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "x-api-token", TokenApiTelegram
    http.Send corpoJSON
    
    If http.Status <> 200 Then EnviarNotificacaoErro = False Else EnviarNotificacaoErro = True

Exit Function
Sair:
        
End Function

Private Function EscapeJSON(ByVal Texto As String) As String
    
Dim resultado As String
    
    resultado = Replace(Texto, "\", "\\")
    resultado = Replace(resultado, """", "\""")
    resultado = Replace(resultado, vbCrLf, "\n")
    resultado = Replace(resultado, vbCr, "\n")
    resultado = Replace(resultado, vbLf, "\n")
    resultado = Replace(resultado, vbTab, "\t")
    
    EscapeJSON = resultado
    
End Function

Private Function FormatarNotificacaoHTML() As String
    
Dim HTML As String
    
    With infNotificacao
        
        .DataHora = Now()
        .versao = FuncoesControlDocs.ExtrairVersaoProjeto
        .EMAIL = FuncoesControlDocs.ObterEmailAssinante
        .Uuid = FuncoesControlDocs.ObterUuidComputador()
        .NomePC = VBA.Environ("COMPUTERNAME")
        
        HTML = "<b>ERRO NO CONTROLDOCS</b>" & vbLf & vbLf
        
        'Dados da Ferramenta
        HTML = HTML & "<b>Data/Hora:</b> " & .DataHora & vbLf
        HTML = HTML & "<b>Versão:</b> " & .versao & vbLf & vbLf
        
        'Dados do Usuário
        HTML = HTML & "<b>--- Dados do usuário ---</b>" & vbLf
        HTML = HTML & "<b>E-mail:</b> " & .EMAIL & vbLf
        HTML = HTML & "<b>Computador:</b> " & .NomePC & vbLf
        HTML = HTML & "<b>UUID:</b> " & .Uuid & vbLf & vbLf
        
        'Detalhes Técnicos
        HTML = HTML & IncluirDetalhesTecnicosHTML()
        
    End With
    
    FormatarNotificacaoHTML = HTML

End Function

Private Function IncluirDetalhesTecnicosHTML() As String

Dim HTML As String
    
    With infNotificacao
        
        HTML = HTML & "<b>--- Detalhes técnicos ---</b>" & vbLf
        If .Funcao <> "" Then HTML = HTML & "<b>Função:</b> " & .Funcao & vbLf
        If .Classe <> "" Then HTML = HTML & "<b>Classe:</b> " & .Classe & vbLf
        If .Modulo <> "" Then HTML = HTML & "<b>Módulo:</b> " & .Modulo & vbLf
        If .ARQUIVO <> "" Then HTML = HTML & "<b>Arquivo:</b> " & .ARQUIVO & vbLf
        If .ConteudoLinha <> "" Then HTML = HTML & "<b>Linha com erro:</b> " & .ConteudoLinha & vbLf
        If .MensagemErro <> "" Then HTML = HTML & "<b>Erro VBA:</b> " & .MensagemErro & vbLf
        If .OBSERVACOES <> "" Then HTML = HTML & "<b>Observações:</b> " & .OBSERVACOES & vbLf
        
    End With
    
    IncluirDetalhesTecnicosHTML = HTML
    
End Function

Public Function NotificarErroInesperado()

Dim Msg As String
    
    Call EnviarNotificacaoErro
    
    Msg = "Ocorreu um erro inesperado ao executar o ControlDocs." & vbCrLf
    Msg = Msg & "Por favor, acione o nosso suporte técnico." & vbCrLf
    
    Call Notificacoes.EnviarNotificacaoErro
    Call Util.MsgAlerta(Msg, "Erro Inesperado")
    
End Function

