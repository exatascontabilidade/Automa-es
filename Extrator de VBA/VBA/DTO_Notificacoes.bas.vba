Attribute VB_Name = "DTO_Notificacoes"
Option Explicit

Public Notificacoes As New ApiNotificacoes
Public infNotificacao As DadosNotificacao

Public Type DadosNotificacao
    
    Funcao As String
    Classe As String
    Modulo As String
    ARQUIVO As String
    MensagemErro As String
    OBSERVACOES As String
    ConteudoLinha As String
    EMAIL As String
    NomePC As String
    Uuid As String
    versao As String
    DataHora As String
    
End Type

