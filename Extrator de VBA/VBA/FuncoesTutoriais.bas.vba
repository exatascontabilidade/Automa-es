Attribute VB_Name = "FuncoesTutoriais"
Option Explicit

Public Function AssistirTutoriais()
            
    Call FuncoesLinks.AbrirUrl(AcessarClub)

End Function

Public Function AssistirTutorial(ByVal Tutorial As String)

    Select Case Tutorial
        
        Case "Autenticacao"
            Call FuncoesLinks.AbrirUrl(videoAutenticarUsuario)
        
        Case "CadContrib"
            ThisWorkbook.FollowHyperlink videoCadastrarContribuinte
            
        Case Else
            ThisWorkbook.FollowHyperlink videoTutorial
            
    End Select

End Function
