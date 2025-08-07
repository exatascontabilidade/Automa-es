Attribute VB_Name = "clsGerenciadorErros"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Sub NotificarErroInesperado(ByVal NomeClasse As String, ByVal NomeMetodo As String, Optional MensagemAdicional As String = "")
    
    With infNotificacao
        
        .Classe = NomeClasse
        .Funcao = NomeMetodo
        .MensagemErro = Err.Number & " - " & Err.Description
        .OBSERVACOES = MensagemAdicional
        
    End With
    
    Call Notificacoes.NotificarErroInesperado
    
End Sub
